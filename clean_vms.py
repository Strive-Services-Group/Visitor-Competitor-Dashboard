#!/usr/bin/env python3
"""
Resident 360 / Visitor Dashboard — daily VMS cleaner.

Scans each project's raw VMS file and rebuilds the dashboard's competitor
dataset (repo file: visitor.xlsx, sheet FINAL).

Rules (per CK, Jul 2026):
  - Keep ONLY these Check In Purposes:
      Laundry, Inspection, Cleaning/Cleaners, Maintenance/Handyman,
      Fit-Out, AMC Contractors  (matched space/case-insensitively)
  - Keep ONLY rows where Check In Type == "Unit Visit".
  - Building/ Unit and Project Name are taken AS-IS from the source file
    (each source Excel maintains its own clean formula) — no derivation here.
  - THE8 file is per-person: collapse duplicate people so one
    (date + purpose + unit + company) counts once. Other files are per-visit.
  - Company Name is upper-cased (cosmetic; dashboard matches case-insensitively
    and strips our own companies Candoo / S&C / Strive on its side).
  - Dates: accept real Excel dates AND text dates in DDMMYYYY[ HH:MM:SS] form
    (e.g. NORTH). Rows with unparseable or implausible dates (year <2024 or
    >2027, e.g. a 4670 typo) are dropped.

Output schema (sheet FINAL), matching the existing repo file:
  Check In Date | Check In Type | Check In Purpose | Company Name |
  Scope of work | Building/ Unit | Project Name

Run:  python3 clean_vms.py
"""
import openpyxl, datetime, sys, os, re, glob

KEEP = {"LAUNDRY", "INSPECTION", "CLEANING/CLEANERS",
        "MAINTENANCE/HANDYMAN", "FIT-OUT", "AMCCONTRACTORS"}
norm = lambda s: str(s or '').upper().replace(' ', '')


def resolve(rel):
    """Find a file/folder under any session mount by its OneDrive-relative path.
    The mount base (/sessions/<name>/mnt) changes every session, so we glob."""
    hits = glob.glob('/sessions/*/mnt/' + rel)
    if hits:
        return hits[0]
    # Fallback: exact path if someone runs it with a fixed mount
    return '/sessions/current/mnt/' + rel


# project, OneDrive-relative path, sheets(None=all), dedupe_per_person
SOURCES = [
 ('BALQIS RESIDENCE',
  "Balqis Security's files - DAILY CONTRACTORS RECORDS/Visitors Details Balqis Residence.xlsx",
  None, False),
 ('THE8',
  "Abdul  Muqeet's files - VMS-DATA FILES/VMS-TH8.xlsx",
  ["VMS-TH8-'26"], True),
 ('NORTH RESIDENCE',
  "VMS DATA SOUTH & NORTH/VISITOR DATA NORTH RESIDENCE - 2026.xlsx",
  ['NORTH RESIDENCE'], False),
 ('SOUTH RESIDENCE',
  "VMS DATA SOUTH & NORTH/VISITOR DATA SOUTH RESIDENCE - 2026.xlsx",
  ['SOUTH RESIDENCE'], False),
]

OUT = resolve("CRM Related/Visitor-Competitor-Dashboard/visitor.xlsx")
HEADER = ['Check In Date', 'Check In Type', 'Check In Purpose', 'Compan