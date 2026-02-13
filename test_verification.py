#!/usr/bin/env python3
"""
Local test of Kodely Hours Verification logic using downloaded CSVs.
Processes Check-ins + Ops Hub data to calculate allowed hours per leader.
"""

import csv
import re
from collections import defaultdict
from datetime import datetime

# ---- FILE PATHS ----
CHECKINS_CSV = "/Users/akashkalimili/Downloads/2026 check-ins - Week of 2_9 (1).csv"
OPS_HUB_CSV = "/Users/akashkalimili/Downloads/Kodely Workshop Ops Hub v2 - Winter_Spring 26 (1).csv"
HOURS_SHEET_CSV = "/Users/akashkalimili/Downloads/Leader Hours Verification - Van's Sheet - 02_05-02_18.csv"

BUFFER_MINUTES = 30
WORKED_STATUSES = ['leader', 'co-lead', 'sub', 'scoot', 'coordinator']

# ---- PARSE TIME ----
def parse_time_to_minutes(time_str):
    """Convert '3:30 PM' to minutes since midnight."""
    if not time_str:
        return None
    time_str = str(time_str).strip().upper()
    match = re.match(r'(\d{1,2}):(\d{2})\s*(AM|PM)?', time_str)
    if not match:
        return None
    hours = int(match.group(1))
    minutes = int(match.group(2))
    period = match.group(3)
    if period == 'PM' and hours != 12:
        hours += 12
    if period == 'AM' and hours == 12:
        hours = 0
    return hours * 60 + minutes

def minutes_to_hm(total_min):
    """Convert minutes to 'Xh Ym' format."""
    h = int(total_min // 60)
    m = int(total_min % 60)
    return f"{h}h {m:02d}m"

def normalize(s):
    """Normalize string for fuzzy matching."""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

# ---- LOAD OPS HUB ----
def load_ops_hub():
    """Load workshop schedules from Ops Hub CSV."""
    workshops = {}
    with open(OPS_HUB_CSV, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            site = (row.get('Site') or '').strip()
            lesson = (row.get('Lesson') or '').strip()
            start_time = (row.get('Start Time') or '').strip()
            end_time = (row.get('End Time') or '').strip()
            setup = (row.get('Setup') or '').strip().lower()
            logistics = (row.get('Logistics') or '').strip().lower()

            # Skip cancelled
            if 'cancelled' in setup or 'cancel' in setup:
                continue
            if 'cancelled' in logistics or 'cancel' in logistics:
                continue

            if not site or not start_time or not end_time:
                continue

            start_min = parse_time_to_minutes(start_time)
            end_min = parse_time_to_minutes(end_time)
            if start_min is None or end_min is None:
                continue

            duration = end_min - start_min
            if duration <= 0:
                continue

            key = normalize(site) + '|' + normalize(lesson)
            workshops[key] = {
                'site': site,
                'lesson': lesson,
                'start_time': start_time,
                'end_time': end_time,
                'duration_min': duration,
                'allowed_min': duration + BUFFER_MINUTES,
            }

            # Also store by just the site for broader matching
            site_key = normalize(site)
            if site_key not in workshops:
                workshops['site:' + site_key + '|' + normalize(lesson)] = workshops[key]

    return workshops

# ---- FIND BEST MATCH ----
def find_workshop(school, workshop, ops_hub):
    """Find the best matching workshop in Ops Hub."""
    # Exact key match
    key = normalize(school) + '|' + normalize(workshop)
    if key in ops_hub:
        return ops_hub[key]

    # Try fuzzy matching
    school_norm = normalize(school)
    workshop_norm = normalize(workshop)

    best_match = None
    best_score = 0

    for k, v in ops_hub.items():
        if k.startswith('site:'):
            continue
        site_norm = normalize(v['site'])
        lesson_norm = normalize(v['lesson'])

        score = 0
        # School/site matching
        if site_norm and school_norm:
            if site_norm == school_norm:
                score += 4
            elif site_norm in school_norm or school_norm in site_norm:
                score += 3
            elif len(site_norm) > 5 and len(school_norm) > 5 and site_norm[:6] == school_norm[:6]:
                score += 2

        # Workshop/lesson matching
        if lesson_norm and workshop_norm:
            if lesson_norm == workshop_norm:
                score += 4
            elif lesson_norm in workshop_norm or workshop_norm in lesson_norm:
                score += 3
            elif len(lesson_norm) > 5 and len(workshop_norm) > 5 and lesson_norm[:6] == workshop_norm[:6]:
                score += 2

        if score > best_score:
            best_score = score
            best_match = v

    return best_match if best_score >= 3 else None

# ---- LOAD CHECK-INS ----
def load_checkins():
    """Load check-in records from CSV."""
    records = []
    with open(CHECKINS_CSV, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)

    # Find header row
    header_row = None
    for i, row in enumerate(rows[:15]):
        row_lower = [str(c).lower().strip() for c in row]
        if any('leader name' in c or 'school' in c for c in row_lower):
            header_row = i
            break

    if header_row is None:
        print("ERROR: Could not find header row in check-ins CSV")
        return records

    headers = [str(c).lower().strip() for c in rows[header_row]]

    def find_col(names):
        for i, h in enumerate(headers):
            for name in names:
                if name in h:
                    return i
        return -1

    col_region = find_col(['region'])
    col_workshop = find_col(['workshop'])
    col_school = find_col(['school', 'site'])
    col_leader = find_col(['leader name'])
    col_date = find_col(['date'])
    col_status = find_col(['status'])

    print(f"Check-ins columns: region={col_region}, workshop={col_workshop}, school={col_school}, leader={col_leader}, date={col_date}, status={col_status}")

    for i in range(header_row + 1, len(rows)):
        row = rows[i]
        if len(row) <= max(col_leader, col_status):
            continue

        leader = str(row[col_leader]).strip() if col_leader >= 0 and col_leader < len(row) else ''
        status = str(row[col_status]).strip().lower() if col_status >= 0 and col_status < len(row) else ''
        workshop = str(row[col_workshop]).strip() if col_workshop >= 0 and col_workshop < len(row) else ''
        school = str(row[col_school]).strip() if col_school >= 0 and col_school < len(row) else ''
        region = str(row[col_region]).strip() if col_region >= 0 and col_region < len(row) else ''
        date_str = str(row[col_date]).strip() if col_date >= 0 and col_date < len(row) else ''

        if not leader or not status:
            continue

        # Check if this is a day separator row (e.g., "TTTuesday", "Wednesday")
        if not workshop and not school:
            continue

        worked = any(ws in status for ws in WORKED_STATUSES)

        records.append({
            'leader': leader,
            'status': status,
            'worked': worked,
            'workshop': workshop,
            'school': school,
            'region': region,
            'date': date_str,
        })

    return records

# ---- LOAD HOURS SHEET (leader names) ----
def load_leader_names():
    """Load known leader names from the Hours Verification sheet."""
    names = set()
    with open(HOURS_SHEET_CSV, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for row in reader:
            if row and row[0].strip() and not any(kw in row[0].lower() for kw in ['leader name', 'bay area', 'sacramento', 'during each', 'log the', 'record the', 'note the', 'how to', 'total hours', 'only mark']):
                name = row[0].strip()
                if len(name) > 2 and not name.startswith(','):
                    names.add(name)
    return names

# ---- MAIN ----
def main():
    print("=" * 70)
    print("KODELY HOURS VERIFICATION - LOCAL TEST")
    print("=" * 70)
    print()

    # Load data
    print("Loading Ops Hub...")
    ops_hub = load_ops_hub()
    print(f"  -> {len(ops_hub)} workshop entries loaded")
    print()

    print("Loading Check-ins...")
    checkins = load_checkins()
    print(f"  -> {len(checkins)} check-in records loaded")
    print()

    print("Loading leader names from Hours Verification sheet...")
    known_leaders = load_leader_names()
    print(f"  -> {len(known_leaders)} known leaders")
    print()

    # Calculate allowed hours per leader
    leader_data = defaultdict(lambda: {
        'sessions': [],
        'total_allowed_min': 0,
        'total_worked': 0,
        'total_absent': 0,
        'unmatched': 0,
    })

    matched_count = 0
    unmatched_count = 0
    unmatched_workshops = set()

    for rec in checkins:
        name = rec['leader']
        ld = leader_data[name]

        if rec['worked']:
            workshop_info = find_workshop(rec['school'], rec['workshop'], ops_hub)

            if workshop_info:
                duration = workshop_info['duration_min']
                allowed = workshop_info['allowed_min']
                matched_count += 1
                source = f"Ops Hub ({workshop_info['site']} - {workshop_info['lesson']})"
            else:
                duration = 60  # default 1 hour
                allowed = duration + BUFFER_MINUTES
                unmatched_count += 1
                ld['unmatched'] += 1
                source = "DEFAULT (1hr assumed)"
                unmatched_workshops.add(f"{rec['school']} | {rec['workshop']}")

            ld['total_allowed_min'] += allowed
            ld['total_worked'] += 1
            ld['sessions'].append({
                'date': rec['date'],
                'workshop': rec['workshop'],
                'school': rec['school'],
                'status': rec['status'],
                'duration_min': duration,
                'allowed_min': allowed,
                'source': source,
            })
        else:
            ld['total_absent'] += 1
            ld['sessions'].append({
                'date': rec['date'],
                'workshop': rec['workshop'],
                'school': rec['school'],
                'status': rec['status'],
                'duration_min': 0,
                'allowed_min': 0,
                'source': 'N/A (not worked)',
            })

    # ---- PRINT REPORT ----
    print("=" * 70)
    print("HOURS VERIFICATION REPORT - Week of Feb 9, 2026")
    print("=" * 70)
    print(f"Match rate: {matched_count} matched / {unmatched_count} unmatched out of {matched_count + unmatched_count} worked sessions")
    print()

    # Summary table
    print(f"{'Leader Name':<35} {'Sessions':>8} {'Absent':>7} {'Allowed Hours':>14} {'Allowed (h:m)':>14}")
    print("-" * 80)

    sorted_leaders = sorted(leader_data.items(), key=lambda x: x[1]['total_allowed_min'], reverse=True)

    total_leaders_with_work = 0
    for name, data in sorted_leaders:
        if data['total_worked'] == 0:
            continue
        total_leaders_with_work += 1
        total_hrs = round(data['total_allowed_min'] / 60, 2)
        formatted = minutes_to_hm(data['total_allowed_min'])
        flag = " *UNMATCHED*" if data['unmatched'] > 0 else ""
        print(f"{name:<35} {data['total_worked']:>8} {data['total_absent']:>7} {total_hrs:>14.2f} {formatted:>14}{flag}")

    print("-" * 80)
    print(f"Total leaders who worked: {total_leaders_with_work}")
    print()

    # Top 20 detailed breakdown
    print("=" * 70)
    print("DETAILED BREAKDOWN (Top 20 by hours)")
    print("=" * 70)
    count = 0
    for name, data in sorted_leaders:
        if data['total_worked'] == 0:
            continue
        count += 1
        if count > 20:
            break

        total_hrs = round(data['total_allowed_min'] / 60, 2)
        print(f"\n{name} — {data['total_worked']} sessions, {minutes_to_hm(data['total_allowed_min'])} allowed ({total_hrs}h)")
        for s in data['sessions']:
            status_icon = "✓" if s['allowed_min'] > 0 else "✗"
            print(f"  {status_icon} {s['date']:>8} | {s['workshop']:<40} @ {s['school']:<30} | {s['status']:<12} | {s['allowed_min']}min | {s['source']}")

    # Unmatched workshops
    if unmatched_workshops:
        print()
        print("=" * 70)
        print(f"UNMATCHED WORKSHOPS ({len(unmatched_workshops)} unique)")
        print("These check-in entries couldn't be matched to the Ops Hub.")
        print("Using default 1hr duration + 30min buffer for these.")
        print("=" * 70)
        for w in sorted(unmatched_workshops):
            print(f"  - {w}")

    # Leaders in Hours Verification sheet but NOT in check-ins
    checkin_leaders = set(leader_data.keys())
    missing_from_checkins = set()
    for known in known_leaders:
        known_norm = normalize(known)
        found = False
        for cl in checkin_leaders:
            if normalize(cl) == known_norm:
                found = True
                break
        if not found:
            missing_from_checkins.add(known)

    if missing_from_checkins:
        print()
        print("=" * 70)
        print(f"LEADERS ON VERIFICATION SHEET BUT NOT IN CHECK-INS ({len(missing_from_checkins)})")
        print("These leaders are listed in the Hours Verification sheet but had")
        print("no check-in records this week (may not have been scheduled).")
        print("=" * 70)
        for name in sorted(list(missing_from_checkins))[:30]:
            print(f"  - {name}")
        if len(missing_from_checkins) > 30:
            print(f"  ... and {len(missing_from_checkins) - 30} more")

    # Save CSV output
    output_csv = "/Users/akashkalimili/kodely-hours-verification/verification_report.csv"
    with open(output_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Leader Name', 'Sessions Worked', 'Sessions Absent', 'Total Allowed Hours', 'Allowed (h:mm)', 'Unmatched Sessions', 'Session Details'])
        for name, data in sorted_leaders:
            if data['total_worked'] == 0 and data['total_absent'] == 0:
                continue
            total_hrs = round(data['total_allowed_min'] / 60, 2)
            formatted = minutes_to_hm(data['total_allowed_min'])
            details = '; '.join([
                f"{s['date']}: {s['workshop']} @ {s['school']} ({s['status']}) -> {s['allowed_min']}min"
                for s in data['sessions']
            ])
            writer.writerow([name, data['total_worked'], data['total_absent'], total_hrs, formatted, data['unmatched'], details])

    print()
    print(f"CSV report saved to: {output_csv}")
    print()
    print("Done!")

if __name__ == '__main__':
    main()
