#!/usr/bin/env python3
"""Shapwick Bunny Hop results processor.

Input: .xlsx sheet with RaceNumber, Time, Firstname, Lastname, Gender,
AgeCategory, ClubName (optional). An existing RacePosition column is ignored.
Output: one audited .xlsx workbook. Does not execute any input workbook macros.
"""
from __future__ import annotations

import argparse
import csv
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, time, timedelta, date
from difflib import SequenceMatcher
import json
import math
from pathlib import Path
import re
import sys
import unicodedata
import zipfile
from typing import Any

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.datetime import from_excel
except ImportError:
    print('The openpyxl package is required. Start with START_BUNNY_HOP.bat to install the required libraries.', file=sys.stderr)
    raise

APP_DIR = Path(__file__).resolve().parent
VERSION = '1.4.3'
DEFAULT_CATEGORIES = ['SM','SF','FV40','FV50','FV60','FV70','MV40','MV50','MV60','MV70','MV80']
COLORS = {'navy':'17324D','teal':'177E89','light':'EAF3F5','gold':'FFF1BC',
          'silver':'E7EBF0','bronze':'F6DFCC','white':'FFFFFF','grey':'576579'}
MANDATORY = ('RaceNumber','Time','Firstname','Lastname','Gender','AgeCategory')
COL_ALIASES = {
    'racenumber':'RaceNumber','raceno':'RaceNumber','bib':'RaceNumber','bibnumber':'RaceNumber',
    'firstname':'Firstname','forename':'Firstname','lastname':'Lastname','surname':'Lastname',
    'agecategory':'AgeCategory','agegroup':'AgeCategory','category':'AgeCategory',
    'clubname':'ClubName','club':'ClubName','gender':'Gender','sex':'Gender',
    'time':'Time','finishtime':'Time',
}

@dataclass
class Runner:
    row: int
    bib: str
    first: str
    last: str
    sex: str
    age: str
    club: str
    seconds: int | None
    overall: int = 0
    sex_pos: int = 0
    age_pos: int = 0
    club_pos: int = 0
    club_sex_pos: int = 0
    overall_award: str = ''
    age_award: str = ''

    @property
    def name(self):
        return f'{self.first} {self.last}'.strip()

    @property
    def display_time(self):
        if self.seconds is None:
            return ''
        minutes, sec = divmod(self.seconds, 60)
        return f'{minutes}:{sec:02d}'


def clean(value: Any) -> str:
    """Safe display cleaning: Unicode spaces, invisible separators and repeated spaces.

    Do not title-case names or remove apostrophes/hyphens: legitimate spelling matters.
    """
    if value is None:
        return ''
    result = unicodedata.normalize('NFKC', str(value))
    for char in ('\u200b', '\u200c', '\u200d', '\ufeff', '\u2060'):
        result = result.replace(char, '')
    return ' '.join(result.split())


def key(value: Any) -> str:
    """Match headings/known codes, ignoring case, whitespace and punctuation."""
    normalized = unicodedata.normalize('NFKD', clean(value)).casefold()
    normalized = ''.join(c for c in normalized if not unicodedata.combining(c))
    return ''.join(c for c in normalized if c.isascii() and c.isalnum())


def normalise_bib(value: Any) -> str:
    """Only unambiguous numeric race numbers: 42, #42, No. 42, and 42.0."""
    raw = clean(value)
    m = re.fullmatch(r'(?:(?:no\.?|number|bib)\s*#?\s*|#\s*)?(\d+)(?:\.0+)?', raw, flags=re.I)
    return m.group(1) if m else raw


def normalise_sex(value: Any) -> str:
    raw = clean(value)
    variants = {
        'm':'M','male':'M','man':'M','men':'M',
        'f':'F','female':'F','woman':'F','women':'F',
    }
    return variants.get(key(raw), raw.upper())


def normalise_age(value: Any, categories: list[str]) -> str:
    raw = clean(value)
    known = {key(c): c for c in categories}
    if 'SM' in categories:
        known.update({'seniormale':'SM', 'malesenior':'SM', 'seniormen':'SM'})
    if 'SF' in categories:
        known.update({'seniorfemale':'SF', 'femalesenior':'SF', 'seniorwomen':'SF'})
    return known.get(key(raw), raw.upper())


def category_sex(category):
    return category[1:] if category in ('SM', 'SF') else category[:1]


def is_blank_club(value: Any) -> bool:
    return key(value) in ('', 'none', 'nan', 'na', 'unaffiliated', 'noclub',
                          'independent', 'notaffiliated', 'noaffiliation', 'nil')


def parse_seconds(value: Any) -> int | None:
    """Original results use MM.SS: numeric 44.1 means 44:10, not 44.1 minutes."""
    if value is None or str(value).strip() == '':
        return None
    if isinstance(value, timedelta):
        sec = value.total_seconds()
        return round(sec) if sec >= 0 else None
    if isinstance(value, datetime):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, (float, int)) and not isinstance(value, bool):
        if not math.isfinite(value) or value < 0:
            return None
        if isinstance(value, float) and 0 < value < 1:
            # A small numeric fraction is usually a native Excel time (fraction of a day).
            # Source data for these two years uses MM.SS values above 1 instead.
            return round(value * 86400)
        s = str(value)
    else:
        s = re.sub(r'\s*([:.])\s*', r'\1', clean(value))
    if ':' in s:
        parts = s.split(':')
        if len(parts) not in (2,3):
            return None
        try:
            vals = [int(x) for x in parts]
        except ValueError:
            return None
        if any(n < 0 for n in vals):
            return None
        if len(vals) == 2:
            m, sec = vals
            return m * 60 + sec if sec < 60 else None
        h, m, sec = vals
        return h * 3600 + m * 60 + sec if m < 60 and sec < 60 else None
    if re.fullmatch(r'\d+\.\d+', s):
        m, sec = s.split('.', 1)
        sec = (sec + '0') if len(sec) == 1 else sec
        # A value with >2 fractional digits is likely damaged float formatting;
        # don't silently reinterpret it.
        if len(sec) > 2 or int(sec) >= 60:
            return None
        return int(m) * 60 + int(sec)
    if re.fullmatch(r'\d+', s):
        return int(s) * 60
    return None


def load_config(config_dir: Path):
    with (config_dir / 'settings.json').open(encoding='utf-8') as f:
        settings = json.load(f)
    with (config_dir / 'clubs_canonical.json').open(encoding='utf-8') as f:
        clubs_cfg = json.load(f)
    cats = [clean(c).upper() for c in settings.get('categories', DEFAULT_CATEGORIES)]
    if not cats or len(set(cats)) != len(cats):
        raise ValueError('Configuration must have unique age categories.')
    canonical = {}
    for group in ('canonical', 'aliases'):
        for raw, display in clubs_cfg.get(group, {}).items():
            normalized = key(raw)
            display = clean(display)
            if not normalized or not display:
                raise ValueError(f'Blank club alias or name in {group}: {raw!r}')
            if normalized in canonical and canonical[normalized] != display:
                raise ValueError(f'Conflicting club aliases for {raw!r}: {canonical[normalized]!r} vs {display!r}')
            canonical[normalized] = display
            # Canonical display names must themselves be accepted, including punctuation variants.
            target = key(display)
            if target in canonical and canonical[target] != display:
                raise ValueError(f'Conflicting canonical club names: {display!r}')
            canonical[target] = display
    return cats, canonical, settings


def read_input(file: Path):
    # read_only=True prevents huge style-only Excel ranges from being loaded into RAM.
    book = load_workbook(file, read_only=True, data_only=True)
    try:
        # Historic 2025 workbooks start with the private FOS export; the
        # established results-entry tab is later in the workbook.
        sheet = next((s for s in book.worksheets if key(s.title) in
                      ('resultsentry','allrunnersresultsinput')), book.worksheets[0])
        rows = sheet.iter_rows(values_only=True)
        headings = next(rows, None)
        if not headings:
            raise ValueError('Input first sheet has no header row.')
        lookup = {}
        for idx, raw in enumerate(headings):
            col = COL_ALIASES.get(key(raw))
            if col:
                if col in lookup:
                    # Historical cleaned results include both authoritative
                    # ClubName and a derived Club column. Prefer ClubName.
                    if col == 'ClubName' and {key(raw), key(headings[lookup[col]])} == {'club', 'clubname'}:
                        if key(raw) == 'clubname':
                            lookup[col] = idx
                        continue
                    raise ValueError(f'Duplicate column after normalisation: {col}')
                lookup[col] = idx
        missing = [col for col in MANDATORY if col not in lookup]
        if missing:
            raise ValueError(f'Missing required columns in results worksheet "{sheet.title}": ' + ', '.join(missing))
        data = []
        for num, values in enumerate(rows, start=2):
            if not any(v is not None and clean(v) for v in values):
                continue
            data.append((num, {name: values[i] if i < len(values) else None
                               for name, i in lookup.items()}))
        return data, sheet.title
    finally:
        book.close()


def transform(input_file: Path, config_dir: Path):
    categories, club_map, settings = load_config(config_dir)
    raw_rows, source_sheet = read_input(input_file)
    issues = []
    def add(severity, code, row, bib, detail):
        issues.append((severity, code, row, bib, detail))

    normalisations = []
    unknown_clubs = {}   # Unrecognised club names: only case/space/punctuation variants coalesce.
    runners = []

    def record_fix(row, bib, field, original, used, reason):
        previous = clean(original)
        if previous != used:
            normalisations.append((row, bib, field, previous, used, reason))

    for row, d in raw_rows:
        bib = normalise_bib(d.get('RaceNumber'))
        first, last = clean(d.get('Firstname')), clean(d.get('Lastname'))
        sex = normalise_sex(d.get('Gender'))
        age = normalise_age(d.get('AgeCategory'), categories)
        raw_club = clean(d.get('ClubName'))
        if is_blank_club(raw_club):
            club = ''
        elif key(raw_club) in club_map:
            club = club_map[key(raw_club)]
        else:
            # Different case/punctuation/spacing for the exact same unknown name
            # can be grouped; never silently merge near-spellings.
            club = unknown_clubs.setdefault(key(raw_club), raw_club)
        seconds = parse_seconds(d.get('Time'))
        runner = Runner(row, bib, first, last, sex, age, club, seconds)
        runners.append(runner)
        record_fix(row, bib, 'RaceNumber', d.get('RaceNumber'), bib, 'Numeric prefix / formatting removed')
        record_fix(row, bib, 'Firstname', d.get('Firstname'), first, 'Spacing / Unicode cleaned')
        record_fix(row, bib, 'Lastname', d.get('Lastname'), last, 'Spacing / Unicode cleaned')
        record_fix(row, bib, 'Gender', d.get('Gender'), sex, 'Case, abbreviation or punctuation normalised')
        record_fix(row, bib, 'AgeCategory', d.get('AgeCategory'), age, 'Case, abbreviation or punctuation normalised')
        if raw_club and not is_blank_club(raw_club):
            record_fix(row, bib, 'ClubName', d.get('ClubName'), club,
                       'Known club alias or matching punctuation/capitalisation')
        if not bib and seconds is not None:
            add('ERROR','Blank finisher race number',row,bib,'A recorded finisher requires a race number.')
        elif bib and not re.fullmatch(r'\d+', bib):
            add('ERROR','Invalid race number',row,bib, 'Expected a whole-number bib; check source entry.')
        if not first or not last:
            add('ERROR','Incomplete name',row,bib,'First and last names are required.')
        if seconds is None and clean(d.get('Time')):
            add('ERROR','Invalid/missing time',row,bib,f'Time value: {d.get("Time")!r}. Excluded from finishing positions.')
        # Blank times are normal for DNS/DNF entrants, who are excluded from placings.
        if sex not in ('M','F'):
            add('ERROR','Unrecognised sex',row,bib,f'Value: {d.get("Gender")!r}.')
        if age not in categories:
            add('ERROR','Unrecognised age category',row,bib,f'Value: {d.get("AgeCategory")!r}.')
        elif category_sex(age) != sex:
            add('ERROR','Category/sex mismatch',row,bib,f'{age} conflicts with sex {sex}.')
        if raw_club:
            if re.fullmatch(r'\d+(?:\.0+)?', raw_club):
                add('REVIEW','Numeric club code',row,bib,f'Club field contains {raw_club!r}; check the original entry system.')
            elif not is_blank_club(raw_club) and key(raw_club) not in club_map:
                suggestions = [(SequenceMatcher(None, key(raw_club), alias).ratio(), name)
                               for alias,name in club_map.items() if alias]
                if suggestions:
                    score,name = max(suggestions, key=lambda x:x[0])
                    if score >= .86 and key(raw_club) != key(name):
                        add('REVIEW','Possible club spelling',row,bib,
                            f'{raw_club!r} may mean {name!r}; not automatically changed.')
    counts = Counter(str(int(r.bib)) for r in runners if re.fullmatch(r'\d+',r.bib))
    for r in runners:
        if re.fullmatch(r'\d+',r.bib) and counts[str(int(r.bib))] > 1:
            add('ERROR','Duplicate race number',r.row,r.bib,'Same numeric bib appears more than once (leading zeros ignored).')
    # A repeated name with exactly the same time and category could be a duplicate row.
    possible_duplicates = defaultdict(list)
    for r in runners:
        if r.first and r.last and r.seconds is not None:
            possible_duplicates[(key(r.first), key(r.last),r.sex,r.age,r.seconds)].append(r)
    for entries in possible_duplicates.values():
        if len(entries) > 1:
            for r in entries:
                add('REVIEW','Possible duplicate runner',r.row,r.bib,
                    'Same name, sex, age category and time appears more than once; verify both bibs.')
    valid = [r for r in runners if r.seconds is not None]
    if not valid:
        raise ValueError('Input contains no finishers with valid times.')
    # Finishing times are the sole source of all places. Do not use any
    # RacePosition value from historic spreadsheets; HQ Controller enters times.
    # For identical displayed times (not expected for Bunny Hop), retain the
    # original input-row order deterministically; no extra data entry required.
    valid.sort(key=lambda r:(r.seconds,r.row))
    sexes=Counter(); ages=Counter(); club_ranks=Counter(); club_sexes=Counter()
    for n,r in enumerate(valid,1):
        r.overall=n
        sexes[r.sex]+=1; r.sex_pos=sexes[r.sex]
        ages[r.age]+=1; r.age_pos=ages[r.age]
        if r.club:
            club_ranks[r.club]+=1; r.club_pos=club_ranks[r.club]
            club_sexes[(r.club,r.sex)]+=1; r.club_sex_pos=club_sexes[(r.club,r.sex)]
    # Assign overall awards BEFORE determining the age-category recipients.
    overall_awards=[]
    for sex in ('M','F'):
        group=[r for r in valid if r.sex==sex]
        for place,r in enumerate(group[:3],1):
            r.overall_award=f'{sex} {place}'
            overall_awards.append((sex,place,r))
    age_awards=[]
    for category in categories:
        eligible=[r for r in valid if r.age==category and r.sex==category_sex(category) and not r.overall_award]
        for place,r in enumerate(eligible[:3],1):
            r.age_award=f'{category} {place}'
            age_awards.append((category, place, r))
    return runners, valid, overall_awards, age_awards, issues, categories, source_sheet, normalisations



def verify_calculations(processed) -> None:
    """Internal cross-checks: fail closed if grouping/ranking/awards disagree.

    These checks validate the processor's arithmetic and list assignments,
    not the accuracy of human-entered bibs or finish times.
    """
    runners, valid, overall_awards, age_awards, _issues, categories, *_ = processed
    def fail(description):
        raise RuntimeError('Internal results calculation check failed: ' + description)

    if len(valid) != sum(r.seconds is not None for r in runners):
        fail('finisher count does not match input time records')
    if len(set(id(r) for r in valid)) != len(valid):
        fail('a finisher appears more than once')
    if set(map(id, valid)) != {id(r) for r in runners if r.seconds is not None}:
        fail('finisher records do not match input records')
    if any(a.seconds > b.seconds for a,b in zip(valid, valid[1:])):
        fail('finishers are not in ascending time order')
    if [r.overall for r in valid] != list(range(1,len(valid)+1)):
        fail('overall places are not consecutive')

    sex_count=Counter(); age_count=Counter(); club_count=Counter(); club_sex_count=Counter()
    for r in valid:
        sex_count[r.sex]+=1
        age_count[r.age]+=1
        if r.sex_pos != sex_count[r.sex]:
            fail('sex-specific places are inconsistent')
        if r.age_pos != age_count[r.age]:
            fail('age-group places are inconsistent')
        if r.club:
            club_count[r.club]+=1
            club_sex_count[(r.club,r.sex)]+=1
            if r.club_pos != club_count[r.club] or r.club_sex_pos != club_sex_count[(r.club,r.sex)]:
                fail('club places are inconsistent')
        elif r.club_pos or r.club_sex_pos:
            fail('unaffiliated finisher assigned club place')
    if sum(sex_count.values()) != len(valid) or sum(age_count.values()) != len(valid):
        fail('sex/age counts do not add up')
    if sum(club_count.values()) != sum(bool(r.club) for r in valid):
        fail('club counts do not add up')

    expected_overall=[(sex,n,r) for sex in ('M','F')
                      for n,r in enumerate([r for r in valid if r.sex==sex][:3],1)]
    if [(sex,n,id(r)) for sex,n,r in overall_awards] != [(sex,n,id(r)) for sex,n,r in expected_overall]:
        fail('overall podium does not match first three male/female finishers')
    for r in valid:
        expected=next((f'{sex} {place}' for sex,place,awarded in expected_overall if awarded is r),'')
        if r.overall_award != expected:
            fail('overall award label does not match allocation')

    expected_ages=[(cat,n,r) for cat in categories
                   for n,r in enumerate([r for r in valid if r.age==cat and r.sex==category_sex(cat) and not r.overall_award][:3],1)]
    if [(cat,n,id(r)) for cat,n,r in age_awards] != [(cat,n,id(r)) for cat,n,r in expected_ages]:
        fail('age awards do not match eligible category finishing order')
    for r in valid:
        expected=next((f'{cat} {place}' for cat,place,awarded in expected_ages if awarded is r),'')
        if r.age_award != expected:
            fail('age award label does not match allocation')
    awarded=[id(r) for _,_,r in overall_awards+age_awards]
    if len(awarded)!=len(set(awarded)):
        fail('runner received more than one individual award')


def excel_safe(value):
    """Keep untrusted cell text as literal data, never an Excel formula."""
    if not isinstance(value,str):
        return value
    if value[:1] in ('=','+','-','@'):
        return "'" + value
    return value


def add_sheet(book, name, headings, entries, subheading=None):
    sheet=book.create_sheet(name)
    sheet.sheet_view.showGridLines = False
    if subheading:
        sheet.append([subheading])
        sheet.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(headings))
        cell=sheet.cell(1,1)
        cell.fill=PatternFill('solid',fgColor=COLORS['teal'])
        cell.font=Font(name='Aptos',size=12,bold=True,color=COLORS['white'])
        cell.alignment=Alignment(vertical='center')
        sheet.row_dimensions[1].height=25
    top=2 if subheading else 1
    for c,label in enumerate(headings,1):
        cell=sheet.cell(top,c, label)
        cell.fill=PatternFill('solid',fgColor=COLORS['navy'])
        cell.font=Font(name='Aptos',size=10,bold=True,color=COLORS['white'])
        cell.alignment=Alignment(vertical='center',wrap_text=True)
    sheet.row_dimensions[top].height=25
    max_width=[len(str(h)) for h in headings]
    for ri,entry in enumerate(entries,top+1):
        for ci,value in enumerate(entry,1):
            val=excel_safe(value)
            cell=sheet.cell(ri,ci,val)
            cell.font=Font(name='Aptos',size=10,color=COLORS['navy'])
            cell.alignment=Alignment(vertical='center')
            if ri % 2 == 0:
                cell.fill=PatternFill('solid',fgColor='F3F7F9')
            max_width[ci-1]=max(max_width[ci-1], len(str(val)) if val is not None else 0)
        sheet.row_dimensions[ri].height=18
    for idx,w in enumerate(max_width,1):
        sheet.column_dimensions[get_column_letter(idx)].width=min(max(w+3,12),42)
    sheet.freeze_panes=f'A{top+1}'
    if entries:
        sheet.auto_filter.ref=f'A{top}:{get_column_letter(len(headings))}{top+len(entries)}'
    sheet.sheet_properties.pageSetUpPr.fitToPage=True
    return sheet


def unique_sheet_name(club, reserved):
    base=re.sub(r'[\\/*?:\[\]]','',club).strip().strip("'") or 'Club'
    base=base[:31]
    name=base; num=2
    while name.lower() in reserved:
        suffix=f' ({num})'
        name=base[:31-len(suffix)]+suffix
        num+=1
    reserved.add(name.lower())
    return name


def create_output(output: Path, source: Path, year: str, processed):
    runners,valid,overall_awards,age_awards,issues,categories,source_sheet,normalisations=processed
    book=Workbook()
    book.remove(book.active)
    counts=Counter(r.sex for r in valid)
    affiliated=[r for r in valid if r.club]
    club_counts=Counter(r.club for r in affiliated)
    missing=sum(1 for r in runners if r.seconds is None)
    warns=sum(1 for x in issues if x[0]=='REVIEW')
    errs=sum(1 for x in issues if x[0]=='ERROR')
    approval_status=('HOLD: correct ERROR checks before publishing' if errs else
                     'REVIEW REQUIRED before publishing' if warns else
                     'READY FOR MANUAL SIGN-OFF')
    add_sheet(book,'Start Here',['Item','Value'],[
        ('Event','Shapwick Bunny Hop '+year),('Source file',source.name),
        ('Input worksheet',source_sheet),('Finishers included',len(valid)),
        ('Male finishers',counts['M']),('Female finishers',counts['F']),
        ('Unaffiliated finishers',sum(1 for r in valid if not r.club)),
        ('Named clubs',len(club_counts)),('Overall individual awards',len(overall_awards)),
        ('Age-group individual awards',len(age_awards)),
        ('Publication status',approval_status),('Calculation self-checks','PASS'),('Error checks',errs),('Review checks',warns),
        ('Automatic formatting fixes',len(normalisations)),('Entrants without finishing time (DNS/DNF or not entered)',missing),
        ('Award rule','Top 3 overall male/female excluded from age-group awards'),
        ('Official results','Overall and age-category positions remain unchanged when awards pass down'),
        ('Source safety','Original input was read only; no macros executed'),
        ('Next step','Review Data Checks, then approve the awards before publication')
    ],'Read this first • review checks before publishing or distributing prizes')
    def runner_common(r):
        return (r.overall,r.bib,r.name,r.sex,r.age,r.age_pos,r.club,r.display_time)
    overall_rows=[('Male' if sex=='M' else 'Female',place,r.overall,r.bib,r.name,r.age,r.age_pos,r.club,r.display_time)
                  for sex,place,r in overall_awards]
    age_rows=[(cat,place,r.age_pos,r.overall,r.bib,r.name,r.sex,r.club,r.display_time)
              for cat,place,r in age_awards]
    overall_sheet=add_sheet(book,'Overall Awards',
        ['Sex','Award Place','Overall Place','Race No','Name','Age Group','True Age Place','Club','Time'],
        overall_rows, 'Overall podiums • 1st–3rd male and female')
    age_sheet=add_sheet(book,'Age Group Awards',
        ['Age Group','Award Place','True Age Place','Overall Place','Race No','Name','Sex','Club','Time'],
        age_rows,'Age-group podiums • overall award winners excluded; full rankings unchanged')
    for sheet,award_col in ((overall_sheet,2),(age_sheet,2)):
        for row in range(3,sheet.max_row+1):
            place=sheet.cell(row,award_col).value
            if place in (1,2,3):
                fill={1:COLORS['gold'],2:COLORS['silver'],3:COLORS['bronze']}[place]
                for cell in sheet[row]:
                    cell.fill=PatternFill('solid',fgColor=fill)
    add_sheet(book,'Non Finishers', ['Input Excel Row','Race No','Name','Sex','Age Group','Club','Time'],
              [(r.row,r.bib,r.name,r.sex,r.age,r.club,'') for r in runners if r.seconds is None],
              'Blank times are acceptable for DNS/DNF; invalid populated times are flagged in Data Checks')
    add_sheet(book,'All Runners',
        ['Overall Place','Race No','Name','Sex','Age Group','Age Place','Club','Time','Overall Award','Age Award'],
        [runner_common(r)+(r.overall_award,r.age_award) for r in valid])
    for sex,title in (('M','Overall Men'),('F','Overall Women')):
        add_sheet(book,title,['Sex Place','Overall Place','Race No','Name','Age Group','Age Place','Club','Time'],
                  [(r.sex_pos,r.overall,r.bib,r.name,r.age,r.age_pos,r.club,r.display_time)
                   for r in valid if r.sex==sex])
    cat_order={cat:i for i,cat in enumerate(categories)}
    age_sorted=sorted(valid,key=lambda r:(cat_order.get(r.age,len(categories)),r.age_pos))
    add_sheet(book,'Age Groups Full',
        ['Age Group','True Age Place','Overall Place','Race No','Name','Sex','Club','Time','Award Received'],
        [(r.age,r.age_pos,r.overall,r.bib,r.name,r.sex,r.club,r.display_time,r.age_award or r.overall_award)
         for r in age_sorted])
    summary=[]
    for club,num in sorted(club_counts.items(),key=lambda item:(-item[1],item[0].lower())):
        sub=[r for r in valid if r.club==club]
        summary.append((club,num,sum(r.sex=='M' for r in sub),sum(r.sex=='F' for r in sub),
                        sub[0].name,sub[0].display_time,sub[0].overall))
    add_sheet(book,'Club Summary',
        ['Club','Finishers','Men','Women','Fastest Finisher','Fastest Time','Best Overall Place'], summary,
        'Affiliated runners only • blank clubs are normal and not reported as errors')
    club_full=sorted(affiliated,key=lambda r:(r.club.lower(),r.club_pos))
    club_cols=['Club','Club Place','Sex Club Place','Overall Place','Race No','Name','Sex','Age Group','Age Place','Time']
    def club_row(r):
        return (r.club,r.club_pos,r.club_sex_pos,r.overall,r.bib,r.name,r.sex,r.age,r.age_pos,r.display_time)
    add_sheet(book,'Club Results Full',club_cols,[club_row(r) for r in club_full])
    sh=[r for r in valid if 'shapwick' in r.club.lower()]
    st=[r for r in valid if r.club.lower()=='street striders']
    add_sheet(book,'Shapwick',club_cols,[club_row(r) for r in sh],
              'Shapwick-affiliated runners, in overall finishing order')
    add_sheet(book,'Street Striders',club_cols,[club_row(r) for r in st],
              'Street Striders, in overall finishing order')
    # Established local-club awards are independent of individual prize pass-down.
    for title, group in (('Shapwick', sh), ('Street Striders', st)):
        seen = set()
        for index, runner in enumerate(group, 3):
            if runner.sex in ('M', 'F') and runner.sex not in seen:
                seen.add(runner.sex)
                for cell in book[title][index]:
                    cell.fill = PatternFill('solid', fgColor='FFFF00')
    # Display race-number and input row; no entrant contact details exposed.
    add_sheet(book,'Data Checks',['Severity','Check','Input Excel Row','Race No','Explanation'],issues,
              'ERROR=HOLD publication • REVIEW=investigate • INFO=not a fault')
    add_sheet(book,'Automatic Fixes',['Input Excel Row','Race No','Field','Original','Used','Reason'],normalisations,
              'Harmless formatting changes made automatically • original input not modified')
    reserved={s.title.lower() for s in book.worksheets}
    for club,num in sorted(club_counts.items(),key=lambda item:(-item[1],item[0].lower())):
        if num < 4 or club.lower() == 'street striders':
            # Street Striders already has a dedicated results worksheet.
            continue
        sheet_name=unique_sheet_name(club,reserved)
        add_sheet(book,sheet_name,club_cols,[club_row(r) for r in valid if r.club==club],
                  f'{club} • {num} finishers')
    book.properties.title=f'Shapwick Bunny Hop {year} classified results and awards'
    book.properties.subject='Overall results, podium awards, age categories and club classifications'
    book.properties.creator='Bunny Hop Results Processor'
    output.parent.mkdir(parents=True,exist_ok=True)
    book.save(output)
    return {'input_records':len(runners),'finishers':len(valid),'awards_overall':len(overall_awards),
            'awards_age':len(age_awards),'errors':errs,'reviews':warns,
            'automatic_fixes':len(normalisations),'calculation_checks':'PASS','publication_status':approval_status,
            'sheets':len(book.worksheets)}


def default_output(input_file:Path, year:str):
    proposed=input_file.parent / f'BunnyHop_{year}_Classified_Results.xlsx'
    if not proposed.exists():
        return proposed
    for n in range(2,1000):
        candidate=input_file.parent / f'BunnyHop_{year}_Classified_Results_v{n}.xlsx'
        if not candidate.exists():
            return candidate
    raise RuntimeError('Too many output versions')


def process(input_file:Path,output:Path|None=None,year:str|None=None,config_dir:Path|None=None):
    input_file=input_file.expanduser().resolve()
    if not input_file.is_file() or input_file.suffix.lower()!='.xlsx':
        raise ValueError('Select an existing .xlsx results input file.')
    if not year:
        found=re.search(r'\b20\d{2}\b',input_file.stem.replace('_',' '))
        year=found.group() if found else str(datetime.now().year)
    if config_dir is None:
        # When packaged as a Windows EXE, allow editable external configuration
        # next to the executable; otherwise use bundled defaults.
        external=(Path(sys.executable).resolve().parent/'config') if getattr(sys,'frozen',False) else APP_DIR/'config'
        config_dir=external if (external/'settings.json').is_file() else APP_DIR/'config'
    output=(output or default_output(input_file,year)).expanduser().resolve()
    if output==input_file:
        raise ValueError('Refusing to overwrite the input workbook.')
    if output.exists():
        raise FileExistsError(f'Output already exists: {output}. Please select another filename.')
    result=transform(input_file,config_dir)
    verify_calculations(result)  # stop if any calculated places, groups or awards disagree
    stats=create_output(output,input_file,year,result)
    return output,stats




# ---------------------------------------------------------------------------
# PRE-RACE ENTRANTS REVIEW
# ---------------------------------------------------------------------------

def parse_dob(value):
    """Parse native Excel dates, date serials and unambiguous common UK date strings."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            got = from_excel(value)
            return got.date() if isinstance(got, datetime) else got if isinstance(got, date) else None
        except (ValueError, OverflowError, TypeError):
            return None
    raw = clean(value)
    if not raw:
        return None
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y'):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def get_age(value):
    raw=clean(value)
    if not raw:
        return None
    try:
        f=float(raw)
        return int(f) if f.is_integer() and 0 <= f <= 115 else None
    except (ValueError, TypeError):
        return None


def easter_monday(year: int) -> date:
    """Western Gregorian Easter Monday; Bunny Hop's normal annual race date."""
    if not 1900 <= int(year) <= 2199:
        raise ValueError('Event year is outside the supported Easter-date range.')
    a=year%19
    b=year//100; c=year%100
    d=b//4; e=b%4
    f=(b+8)//25; g=(b-f+1)//3
    h=(19*a+b-d-g+15)%30
    i=c//4; k=c%4
    l=(32+2*e+2*i-h-k)%7
    m=(a+11*h+22*l)//451
    month=(h+l-7*m+114)//31
    day=(h+l-7*m+114)%31+1
    return date(year,month,day)+timedelta(days=1)


def age_on_date(dob: date, race_date: date) -> int:
    return race_date.year - dob.year - ((race_date.month, race_date.day) < (dob.month, dob.day))


def age_category_from_age(age: int, sex: str, categories: list[str]) -> str:
    """Use configured categories only; oldest category covers all older competitors.

    Actual prize age rules require human approval; never infer from the
    incomplete AgeCategory value in the entry-system export.
    """
    if sex not in ('M','F') or age is None or age < 0:
        return ''
    candidates=[]
    for cat in categories:
        match=re.fullmatch(r'([MF])V(\d+)',cat)
        if match and match.group(1)==sex:
            candidates.append((int(match.group(2)),cat))
    candidates.sort()
    selected='S'+sex
    for minimum,cat in candidates:
        if age >= minimum:
            selected=cat
    return selected if selected in categories else ''


def merge_manual_master(fos_data, master_sheet):
    """Combine the PRIVATE original FOS data with a separately numbered Master list.

    FOS never assigns race numbers here. The manually created Master list is the
    authority for race numbers, corrected club text and award category. Match by
    name+sex only where safe; duplicates are explicitly marked for review.
    """
    headers=next(master_sheet.iter_rows(min_row=1,max_row=1,values_only=True),None)
    if not headers:
        return fos_data
    columns={key(COL_ALIASES.get(key(h),h)):i for i,h in enumerate(headers) if h is not None}
    required={'firstname','lastname','gender','racenumber','agecategory'}
    if not required.issubset(columns):
        return fos_data
    master=[]
    for row_num,values in enumerate(master_sheet.iter_rows(min_row=2,values_only=True),2):
        if not any(v is not None and clean(v) for v in values):
            continue
        d={k:(values[i] if i<len(values) else None) for k,i in columns.items()}
        master.append((row_num,d))
    if not master:
        return fos_data

    def identity(d):
        return (key(d.get('firstname')),key(d.get('lastname')),
                normalise_sex(d.get('gender')))

    fos_by_key=defaultdict(list)
    master_by_key=defaultdict(list)
    for row,d in fos_data:
        fos_by_key[identity(d)].append((row,d))
    for row,d in master:
        master_by_key[identity(d)].append((row,d))

    merged=[]
    for ident,group in master_by_key.items():
        originals=fos_by_key.get(ident,[])
        if len(originals)!=len(group):
            note=('This master-list name/sex has %d matching FOS entry/entries '
                  'but %d master-list record(s); reconcile manually.'
                  % (len(originals),len(group)))
        elif len(group)>1:
            note=('Multiple registrations have the same name and sex. Bib-to-FOS '
                  'entry-ID pairing cannot be verified from this sheet; '
                  'check whether these are genuine separate entrants.')
        else:
            note=''
        for i,(master_row,md) in enumerate(group):
            if i<len(originals):
                original_row,od=originals[i]
                record=od.copy()
                record['_fos_original_row']=original_row
            else:
                record={}
            # Values manually curated in the master are the authoritative
            # race-day roster fields. Blank club is a legitimate value.
            for field in ('firstname','lastname','gender','racenumber',
                          'agecategory','clubname'):
                if field in md:
                    record[field]=md[field]
            record['_from_manual_master']=True
            if note:
                record['_manual_master_issue']=note
            merged.append((master_row,record))
    # Registrations missing from the manually prepared master must be visible.
    for ident,group in fos_by_key.items():
        used=min(len(group),len(master_by_key.get(ident,[])))
        for fos_row,od in group[used:]:
            record=od.copy()
            record['_manual_master_issue']=(
                'FOS entry not matched to any row in the manually numbered Master list. '
                'Check whether the entrant was omitted or withdrawn.')
            merged.append((fos_row,record))
    return merged


def read_participants(file: Path):
    """Read original FOS CSV or Excel master; merge a second Master list sheet when present.

    FOS does not assign race numbers. No workbook code or macros are executed.
    """
    def from_rows(row_iter, sheet_title, allow_master=False):
        iterator=iter(row_iter)
        header=next(iterator,None)
        if not header:
            raise ValueError('Participants workbook has no header row.')
        columns={}
        for i,h in enumerate(header):
            if h is not None:
                columns.setdefault(key(COL_ALIASES.get(key(h),h)),i)
        required=('firstname','lastname','gender')
        missing=[col for col in required if col not in columns]
        if missing:
            raise ValueError('FOS participants export is missing fields: '+', '.join(missing))
        manual = allow_master and 'racenumber' in columns and 'agecategory' in columns and not any(field in columns for field in ('ageonday','ageonraceday','dob'))
        if not manual and not any(field in columns for field in ('ageonday','ageonraceday','dob')):
            raise ValueError('FOS participants export needs DOB, AgeOnDay or Age on Race Day.')
        rows=[]
        for excel_row,values in enumerate(iterator,2):
            if not any(v is not None and clean(v) for v in values):
                continue
            record={k:values[idx] if idx<len(values) else None for k,idx in columns.items()}
            if manual:
                record['_from_manual_master']=True
            rows.append((excel_row,record))
        return rows,sheet_title

    ext=file.suffix.lower()
    if ext=='.csv':
        # FOS's normal participant download: CSV with a heading row.
        # utf-8-sig removes an optional BOM; cp1252 accommodates older Windows exports.
        try:
            raw=file.read_bytes()
            try:
                decoded=raw.decode('utf-8-sig')
            except UnicodeDecodeError:
                decoded=raw.decode('cp1252')
            import io
            with io.StringIO(decoded,newline='') as textfile:
                reader=csv.reader(textfile,delimiter=',',strict=True)
                return from_rows(reader, 'FOS CSV: '+file.name)
        except (csv.Error, UnicodeError) as exc:
            raise ValueError('Cannot read the FOS CSV file. Download it again from FOS; do not edit it in a text editor.') from exc
    if ext not in ('.xlsx','.xls'):
        raise ValueError('Select a FOS .csv export or a prepared .xlsx/.xls master workbook.')
    if ext=='.xlsx' or zipfile.is_zipfile(file):
        # Passing an open file stream also accepts a real .xlsx saved with .xls extension.
        with file.open('rb') as fp:
            book=load_workbook(fp,read_only=True,data_only=True,keep_links=False)
            try:
                sheet=book.worksheets[0]
                fos_data,source=from_rows(sheet.iter_rows(values_only=True),sheet.title,allow_master=True)
                master=next((ws for ws in book.worksheets[1:]
                             if key(ws.title)=='masterlist'),None)
                if master is not None:
                    merged=merge_manual_master(fos_data,master)
                    # Keep explicit provenance so that a human knows which
                    # original worksheet a flagged row refers to.
                    if merged is not fos_data:
                        return merged, master.title+' (matched with FOS data)'
                return fos_data,source
            finally:
                book.close()
    # xlrd 2+ supports genuine binary .xls files, without executing macros.
    with file.open('rb') as fp:
        signature=fp.read(8)
    if signature != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        raise ValueError('This .xls is not a genuine binary Excel workbook. '
                         'Open the FOS download in Excel and use File > Save As > '
                         'Excel Workbook (.xlsx), then select the saved copy.')
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError('Reading genuine .xls files requires xlrd. '
                           'Use the Windows build with xlrd, or Save As .xlsx in Excel.') from exc
    book=xlrd.open_workbook(str(file),on_demand=True)
    try:
        sheet=book.sheet_by_index(0)
        def iter_rows():
            for ri in range(sheet.nrows):
                record=[]
                for ci in range(sheet.ncols):
                    cell=sheet.cell(ri,ci)
                    if cell.ctype==xlrd.XL_CELL_DATE:
                        try:
                            record.append(xlrd.xldate_as_datetime(cell.value,book.datemode))
                        except (ValueError,OverflowError,TypeError):
                            record.append(cell.value)
                    elif cell.ctype in (xlrd.XL_CELL_EMPTY,xlrd.XL_CELL_BLANK):
                        record.append(None)
                    elif cell.ctype==xlrd.XL_CELL_ERROR:
                        record.append('#EXCEL_ERROR')
                    else:
                        record.append(cell.value)
                yield record
        return from_rows(iter_rows(),sheet.name,allow_master=True)
    finally:
        book.release_resources()


def analyse_participants(file:Path,race_date:date|None,config_dir:Path):
    categories,clubs,_=load_config(config_dir)
    data,source_sheet=read_participants(file)
    issues=[];fixes=[];entrants=[]
    def issue(severity,code,excel_row,entry_id,detail):
        issues.append((severity,code,excel_row,entry_id,detail))
    if not race_date:
        issue('REVIEW','Race date not supplied','','',
              'Ages/categories rely on supplied AgeOnDay; enter the exact race date to verify DOB against age and categories.')
    seen_ids=defaultdict(list)
    seen_people=defaultdict(list)
    unknown_clubs={}
    for row,d in data:
        entry_id=clean(d.get('epid'))
        if d.get('_manual_master_issue'):
            issue('REVIEW','Master and FOS reconciliation',row,entry_id,
                  d['_manual_master_issue'])
        first=clean(d.get('firstname'))
        last=clean(d.get('lastname'))
        sex=normalise_sex(d.get('gender'))
        original_club=clean(d.get('clubname'))
        # Exact export uses ClubName; TeamName is NOT the club.
        # If the source has duplicate Club Name headings, the first export field is authoritative.
        club='' if is_blank_club(original_club) else clubs.get(key(original_club), original_club)
        if club and key(original_club) not in clubs:
            club=unknown_clubs.setdefault(key(original_club),club)
        if original_club and not is_blank_club(original_club) and original_club!=club:
            fixes.append((row,entry_id,'ClubName',original_club,club,
                          'Exact known alias, case, spacing or punctuation normalised'))
        dob=parse_dob(d.get('dob'))
        declared=get_age(d.get('ageonday'))
        declared_alt=get_age(d.get('ageonraceday'))
        if clean(d.get('dob')) and dob is None:
            issue('REVIEW','Unrecognised DOB format',row,entry_id,'Cannot interpret date of birth; check the original registration.')
        if dob and (dob.year < 1900 or dob > date.today()):
            issue('ERROR','Impossible date of birth',row,entry_id,'DOB is outside a plausible range.')
            dob=None
        if clean(d.get('ageonday')) and declared is None:
            issue('REVIEW','Invalid AgeOnDay',row,entry_id,'AgeOnDay must be a whole number between 0 and 115.')
        if clean(d.get('ageonraceday')) and declared_alt is None:
            issue('REVIEW','Invalid Age on Race Day',row,entry_id,'Age on Race Day must be a whole number between 0 and 115.')
        if declared is not None and declared_alt is not None and declared!=declared_alt:
            issue('REVIEW','Conflicting supplied ages',row,entry_id,
                  f'AgeOnDay {declared} differs from Age on Race Day {declared_alt}.')
        if dob and race_date and dob > race_date:
            issue('ERROR','DOB after race date',row,entry_id,'DOB is later than selected race date.')
        computed=age_on_date(dob,race_date) if dob and race_date and dob<=race_date else None
        age=computed if computed is not None else declared if declared is not None else declared_alt
        if age is None and not d.get('_from_manual_master'):
            issue('ERROR','Missing usable age',row,entry_id,'Supply a valid race-date age or DOB plus a race date.')
        if computed is not None:
            for title,a in (('AgeOnDay',declared),('Age on Race Day',declared_alt)):
                if a is not None and a!=computed:
                    issue('REVIEW','DOB and supplied age disagree',row,entry_id,
                          f'{title} {a} differs from DOB-derived age {computed}. Using DOB-derived age until checked.')
        if not first or not last:
            issue('ERROR','Missing entrant name',row,entry_id,'First and last names are required. Record omitted from results-entry template.')
        if sex not in ('M','F') and first and last:
            issue('ERROR','Unrecognised sex',row,entry_id,'Cannot assign a sex-specific prize category.')
        if not entry_id and not d.get('_from_manual_master'):
            issue('REVIEW','Missing entry identifier',row,entry_id,
                  'No ep_id supplied; use name and DOB to reconcile the original registration.')
        elif entry_id:
            seen_ids[entry_id].append(row)
        cat=age_category_from_age(age,sex,categories) if age is not None else ''
        if age is not None and not cat and sex in ('M','F'):
            issue('ERROR','No matching age category',row,entry_id,
                  f'Age {age} / sex {sex} does not map to configured award categories.')
        # FOS may contain its OWN event groups, which can be different from
        # Bunny Hop prize categories (e.g. 'Senior male 60' for a 53-year-old).
        # Compare only if the source already identifies an exact BH category.
        fos_category=clean(d.get('agecategory'))
        fos_bh_category=normalise_age(fos_category,categories) if fos_category else ''
        if fos_bh_category in categories and cat and fos_bh_category!=cat:
            issue('REVIEW','Age category differs from derived age',row,entry_id,
                  f'Entered category {fos_category!r} maps to {fos_bh_category}; '
                  f'calculated Bunny Hop group is {cat}. Check race date and prize rules.')
        if d.get('_from_manual_master'):
            cat=fos_bh_category
            if cat not in categories:
                issue('ERROR','Unrecognised master category',row,entry_id,
                      'Correct the AgeCategory in the numbered Master list.')
            elif category_sex(cat)!=sex:
                issue('ERROR','Category/sex mismatch',row,entry_id,
                      f'{cat} conflicts with sex {sex}. Correct the Master list.')
        if age is not None and age<18:
            issue('REVIEW','Junior entrant',row,entry_id,
                  'Age under 18: check the event entry conditions and eligibility; no minimum age assumed by the program.')
        if first and last:
            seen_people[(key(first),key(last),dob.isoformat() if dob else '',sex)].append(row)
        if original_club and not is_blank_club(original_club):
            if re.fullmatch(r'\d+(?:\.0+)?',original_club):
                issue('REVIEW','Numeric club value',row,entry_id,'Club field contains a numeric code; verify actual club.')
            elif key(original_club) not in clubs:
                choices=[(SequenceMatcher(None,key(original_club),alias).ratio(),name)
                         for alias,name in clubs.items() if alias]
                if choices:
                    score,suggestion=max(choices,key=lambda x:x[0])
                    if score>=.86 and key(original_club)!=key(suggestion):
                        issue('REVIEW','Possible club spelling',row,entry_id,
                              f'Unrecognised club {original_club!r}, perhaps {suggestion!r}; not automatically changed.')
        for label,original,used in (('Firstname',d.get('firstname'),first),('Lastname',d.get('lastname'),last),
                                   ('Gender',d.get('gender'),sex)):
            if clean(original)!=used or (original is not None and str(original)!=used):
                fixes.append((row,entry_id,label,clean(original),used,'Harmless space/case normalisation'))
        if age is not None and computed is not None and declared!=computed:
            fixes.append((row,entry_id,'Race-day age',str(declared) if declared is not None else '',
                          str(computed),'Computed from DOB and selected race date; review any discrepancy'))
        bib=normalise_bib(d.get('racenumber'))
        # No bib is NORMAL in a pre-race export; only inspect nonblank values.
        if bib and not re.fullmatch(r'\d+',bib):
            issue('REVIEW','Invalid allocated race number',row,entry_id,'Non-numeric race number if already assigned.')
        if clean(d.get('racenumber')) and bib!=clean(d.get('racenumber')):
            fixes.append((row,entry_id,'RaceNumber',clean(d.get('racenumber')),bib,
                          'Unambiguous numeric formatting normalised'))
        entrants.append({'row':row,'id':entry_id,'first':first,'last':last,'sex':sex,'age':age,
                         'category':cat,'club':club,'bib':bib,'fos_category':fos_category})
    for eid,rows in seen_ids.items():
        if len(rows)>1:
            for row in rows:
                issue('ERROR','Duplicate entry identifier',row,eid,
                      'The same ep_id appears more than once in the participants export.')
    for person,rows in seen_people.items():
        if len(rows)>1:
            for row in rows:
                issue('REVIEW','Possible duplicate entrant',row,'',
                      'Same name, DOB (when available), and sex appears in multiple entries; check manually.')
    bibs=defaultdict(list)
    for entrant in entrants:
        if entrant['bib'] and entrant['bib'].isdigit():
            bibs[str(int(entrant['bib']))].append(entrant)
    for bib,entries in bibs.items():
        if len(entries)>1:
            for entrant in entries:
                issue('ERROR','Duplicate allocated race number',entrant['row'],entrant['id'],
                      f'Race number {bib} appears more than once.')
    # FOS has blank bibs before manual allocation. Warn ONCE if partly numbered;
    # do not flood the report with a warning for every empty bib.
    named=[e for e in entrants if e['first'] and e['last']]
    allocated=[e for e in named if e['bib'] and e['bib'].isdigit()]
    if 0 < len(allocated) < len(named):
        issue('REVIEW','Partially allocated race numbers','','',
              f'{len(allocated)} of {len(named)} named entrants have numeric race numbers; '
              'normal while numbering is in progress, check before printing the final race list.')
    # Intentionally NO checks on the sequence or alphabetical ordering of race
    # numbers. Bib allocation order is an organiser convenience, not a rule.
    # Only uniqueness and validity matter. The report lists entrants in
    # alphabetical order to help race-day collection without renumbering them.
    return entrants,issues,fixes,source_sheet,categories


def available_pair(input_file:Path,year:str):
    for version in range(1,1000):
        postfix='' if version==1 else f'_v{version}'
        review=input_file.parent/f'BunnyHop_{year}_Entry_Check{postfix}.xlsx'
        entry=input_file.parent/f'BunnyHop_{year}_Results_Entry{postfix}.xlsx'
        if not review.exists() and not entry.exists():
            return review,entry
    raise RuntimeError('Too many output versions exist.')


def check_entrants(input_file:Path,year:str,race_date:date|str|None=None,config_dir:Path|None=None):
    """Create PRIVATE validation report and a SAFE minimum-field results-entry file.

    Never edits the source CSV or workbook. Rejects existing output names.
    """
    input_file=Path(input_file).expanduser().resolve()
    if not input_file.is_file() or input_file.suffix.lower() not in ('.csv','.xlsx','.xls'):
        raise ValueError('Select an existing FOS participants .csv file or a prepared .xlsx/.xls master workbook.')
    if not re.fullmatch(r'20\d{2}',str(year)):
        raise ValueError('Event year must be YYYY (e.g. 2027).')
    if isinstance(race_date,str):
        race_date=date.fromisoformat(race_date) if race_date.strip() else None
    if race_date is None:
        race_date=easter_monday(int(year))
    if race_date and str(race_date.year)!=str(year):
        raise ValueError('Selected race date must be in the specified event year.')
    if config_dir is None:
        external=(Path(sys.executable).resolve().parent/'config') if getattr(sys,'frozen',False) else APP_DIR/'config'
        config_dir=external if (external/'settings.json').is_file() else APP_DIR/'config'
    entrants,issues,fixes,source_sheet,categories=analyse_participants(input_file,race_date,config_dir)
    review,entry=available_pair(input_file,year)
    problems=Counter(x[0] for x in issues)
    status=('HOLD — correct errors before race-day use' if problems['ERROR'] else
            'REVIEW — check flagged entries before race-day use' if problems['REVIEW'] else
            'READY FOR HUMAN REVIEW')
    report=Workbook();report.remove(report.active)
    add_sheet(report,'Start Here',['Item','Value'],[
        ('Event',f'Shapwick Bunny Hop {year}'),('Input file',input_file.name),
        ('Race date',str(race_date)+' (Easter Monday by default; can be overridden)'),
        ('Input worksheet',source_sheet),('Entry records',len(entrants)),
        ('Race-number source','Manually prepared Master list when present; NEVER allocated by FOS.'),
        ('Named entrant rows',sum(bool(e['first'] and e['last']) for e in entrants)),
        ('Race numbers allocated',sum(bool(e['bib']) for e in entrants)),
        ('Number allocation','Manual; unique bib per entrant. Any numbering order is acceptable.'),
        ('Errors',problems['ERROR']),('Items to review',problems['REVIEW']),
        ('Automatic fixes / age calculations',len(fixes)),('Status',status),
        ('Important','Blank clubs, race numbers and finish times are normal before the race.'),
        ('Category rule','Numbered Master categories preserved and cross-checked when age is available. Raw FOS categories derived from race-day age and sex.'),
        ('Privacy','No addresses, email, emergency contacts or medical details exported.'),
        ('Action','Resolve Issues to Review, then check and use Results Entry template.')
    ], 'Private pre-race report • resolve ERROR and REVIEW items before race day')
    add_sheet(report,'Issues to Review', ['Severity','Check','Original Excel Row','Entry ID','Explanation'],issues,
              'ERROR=must resolve • REVIEW=human check • INFO=information')
    add_sheet(report,'Clean Entrants',['Source Row','Entry ID','Firstname','Lastname','Sex','Age on Race Day','Age Category','Source Age Group','Club','Race No'],
              [(e['row'],e['id'],e['first'],e['last'],e['sex'],e['age'],e['category'],e['fos_category'],e['club'],e['bib']) for e in entrants],
              'Private roster: NO addresses, contacts, dates of birth or medical data')
    add_sheet(report,'Automatic Fixes',['Source Row','Entry ID','Field','Source Value','Used Value','Reason'],fixes)
    ordered_names=sorted((e for e in entrants if e['first'] and e['last']),
                         key=lambda e:(key(e['last']),key(e['first']),e['row']))
    add_sheet(report,'Race Number Audit',
              ['Alphabetical Order','Source Row','Entry ID','Surname','Firstname','Allocated Race No'],
              [(n,e['row'],e['id'],e['last'],e['first'],e['bib']) for n,e in enumerate(ordered_names,1)],
              'Alphabetical collection list: bib numbers shown as allocated; ordering of bib numbers is NOT checked')
    report.properties.title=f'Bunny Hop {year} entrants validation (private)'
    # Results entry template must have the same eight first-sheet columns the classifier expects.
    # The only extra field is an entry-system ID for reconciliation. No private contact details.
    result=Workbook();result.active.title='Results Entry'
    ws=result.active
    headers=['RaceNumber','Time','Firstname','Lastname','Gender','AgeCategory','ClubName','Entry ID','Source Row']
    ws.append(headers)
    for e in entrants:
        if not e['first'] or not e['last']:
            continue
        ws.append([e['bib'],'',excel_safe(e['first']),excel_safe(e['last']),
                   e['sex'],e['category'],excel_safe(e['club']),excel_safe(e['id']),e['row']])
    ws.freeze_panes='C2';ws.auto_filter.ref=f'A1:I{ws.max_row}';ws.sheet_view.showGridLines=False
    for c,label in enumerate(headers,1):
        cell=ws.cell(1,c)
        cell.fill=PatternFill('solid',fgColor=COLORS['navy'])
        cell.font=Font(name='Aptos',bold=True,color='FFFFFF')
        ws.column_dimensions[get_column_letter(c)].width=min(max(len(label)+3,14),28)
    ws.column_dimensions['C'].width=23;ws.column_dimensions['D'].width=23
    ws.column_dimensions['G'].width=32
    ws.column_dimensions['B'].width=18
    result.properties.title=f'Bunny Hop {year} results entry template'
    result.properties.subject='Minimal results fields; no personal contact or health information'
    # Write to temporary locations first so conversion errors don't leave a partial pair.
    from tempfile import TemporaryDirectory
    with TemporaryDirectory(prefix='bunnyhop_staging_',dir=str(input_file.parent)) as tmp:
        stage_a=Path(tmp)/review.name;stage_b=Path(tmp)/entry.name
        report.save(stage_a);result.save(stage_b)
        if review.exists() or entry.exists():
            raise FileExistsError('Output filename is already in use. Nothing was overwritten.')
        stage_a.replace(review);stage_b.replace(entry)
    return review,entry,{'records':len(entrants),'ready_named':sum(bool(e['first'] and e['last']) for e in entrants),
                         'race_numbers_allocated':sum(bool(e['bib']) for e in entrants),
                         'errors':problems['ERROR'],'reviews':problems['REVIEW'],
                         'fixes':len(fixes),'status':status}

def gui():
    import tkinter as tk
    from tkinter import filedialog,messagebox,ttk
    root=tk.Tk();root.title('Bunny Hop 1.4.3 — Entrants & Results');root.geometry('780x480');root.minsize(740,460)
    var_file=tk.StringVar();var_year=tk.StringVar(value=str(datetime.now().year))
    var_date=tk.StringVar(value=easter_monday(int(var_year.get())).isoformat());var_status=tk.StringVar(value='Choose a participants export or results spreadsheet. Original files remain unchanged.')
    last_auto_date=[var_date.get()]
    def sync_auto_date(*_):
        year=var_year.get().strip()
        if re.fullmatch(r'20\d{2}',year):
            auto=easter_monday(int(year)).isoformat()
            if not var_date.get().strip() or var_date.get()==last_auto_date[0]:
                var_date.set(auto)
            last_auto_date[0]=auto
    var_year.trace_add('write',sync_auto_date)
    frame=ttk.Frame(root,padding=20);frame.pack(fill='both',expand=True)
    ttk.Label(frame,text='Shapwick Bunny Hop — Entrants & Results',font=('Segoe UI',17,'bold')).pack(anchor='w',pady=(0,6))
    ttk.Label(frame,text='Two stages: pre-race checks, then post-race classified results and awards.').pack(anchor='w',pady=(0,14))
    pick=ttk.Frame(frame);pick.pack(fill='x')
    ttk.Entry(pick,textvariable=var_file,width=78).pack(side='left',fill='x',expand=True)
    def browse():
        filename=filedialog.askopenfilename(title='Select participants export OR results workbook',filetypes=[('FOS CSV and Excel workbooks','*.csv *.xlsx *.xls'),('All files','*.*')])
        if filename:
            var_file.set(filename)
            match=re.search(r'20\d{2}',Path(filename).stem)
            if match:var_year.set(match.group())
    ttk.Button(pick,text='Browse…',command=browse).pack(side='left',padx=(8,0))
    line=ttk.Frame(frame);line.pack(anchor='w',pady=(13,8))
    ttk.Label(line,text='Event year:').pack(side='left')
    ttk.Entry(line,textvariable=var_year,width=8).pack(side='left',padx=(7,30))
    ttk.Label(line,text='Easter Monday date (editable):').pack(side='left')
    ttk.Entry(line,textvariable=var_date,width=14).pack(side='left',padx=7)
    ttk.Label(frame,text='PRE-RACE: select the original FOS CSV or your numbered Excel master workbook. Bibs are added manually, not by FOS.',wraplength=680).pack(anchor='w',pady=(4,8))
    def check():
        try:
            year=var_year.get().strip()
            d=var_date.get().strip()
            report,template,stats=check_entrants(Path(var_file.get()),year,d)
            var_status.set(f'{stats["status"]}: {report.name}')
            messagebox.showwarning('Pre-race checks complete',
                 f'Input entries: {stats["records"]}\nNamed entrants: {stats["ready_named"]}\n'
                 f'Errors: {stats["errors"]}; review items: {stats["reviews"]}\n'
                 f'Race numbers already allocated: {stats["race_numbers_allocated"]}\n\n'
                 f'Private report: {report}\nResults-entry sheet: {template}\n\n'
                 'Review Issues to Review before race day. No contact or health fields were exported.')
        except Exception as exc:
            messagebox.showerror('Pre-race check failed',str(exc))
            var_status.set('Pre-race check could not complete; original file unchanged.')
    ttk.Button(frame,text='1. CHECK ENTRANTS (BEFORE RACE)',command=check).pack(anchor='w',pady=(0,13))
    ttk.Label(frame,text='POST-RACE: select the completed results-entry workbook, with finish times and race numbers added.',wraplength=680).pack(anchor='w',pady=(3,7))
    def go():
        try:
            year=var_year.get().strip()
            if not re.fullmatch(r'20\d{2}',year):
                raise ValueError('Enter a valid event year, e.g. 2027.')
            out,summary=process(Path(var_file.get()),year=year)
            msg=(f'Created: {out}\n\nFinishers: {summary["finishers"]}\n'
                 f'Overall prizes: {summary["awards_overall"]}\n'
                 f'Age-group prizes: {summary["awards_age"]}\n'
                 f'Errors: {summary["errors"]}; review flags: {summary["reviews"]}\n'
                 f'Status: {summary["publication_status"]}\n\n'
                 'Check Data Checks and approve the awards before publication.')
            var_status.set(f'{summary["publication_status"]}: {out.name}')
            if summary['errors'] or summary['reviews']:
                messagebox.showwarning('Results created — review required',msg)
            else:
                messagebox.showinfo('Results workbook created',msg)
        except Exception as exc:
            messagebox.showerror('Could not process results',str(exc))
            var_status.set('Results processing failed — original file unchanged.')
    ttk.Button(frame,text='2. PROCESS RESULTS (AFTER RACE)',command=go).pack(anchor='w',pady=(0,12))
    ttk.Label(frame,textvariable=var_status,wraplength=690,foreground='#36576C').pack(anchor='w',pady=(0,10))
    ttk.Label(frame,text='HQ Controller processes results; Race Director approves final awards.').pack(anchor='w')
    root.mainloop()


def main():
    parser=argparse.ArgumentParser(description='Bunny Hop pre-race entrant validation and post-race awards.')
    modes=parser.add_mutually_exclusive_group()
    modes.add_argument('--input',type=Path,help='Post-race results .xlsx; omit for GUI')
    modes.add_argument('--check-entrants',type=Path,help='Pre-race FOS CSV (.csv) or numbered Excel master (.xlsx/.xls)')
    parser.add_argument('--output',type=Path,help='Post-race output .xlsx (must not exist)')
    parser.add_argument('--year',help='Event year, e.g. 2027')
    parser.add_argument('--race-date',help='Optional override YYYY-MM-DD; defaults to Easter Monday in the event year')
    parser.add_argument('--config',type=Path,help='Folder with settings.json and clubs_canonical.json')
    args=parser.parse_args()
    if args.check_entrants:
        if args.output:
            parser.error('--output applies only to post-race results processing')
        year=args.year
        if not year:
            match=re.search(r'20\d{2}',args.check_entrants.name)
            year=match.group() if match else str(datetime.now().year)
        review,entry,stats=check_entrants(args.check_entrants,year,args.race_date,args.config)
        print(json.dumps({'review':str(review),'results_entry':str(entry),**stats},indent=2));return
    if not args.input:
        gui();return
    output,stats=process(args.input,args.output,args.year,args.config)
    print(json.dumps({'output':str(output), **stats},indent=2))

if __name__=='__main__':
    main()
