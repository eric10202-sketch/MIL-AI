/**
 * generate_msp_xml.js
 * Converts Trinity_Project_Schedule.csv → Trinity_Project_Schedule.xml
 * Microsoft Project XML format (schema: http://schemas.microsoft.com/project)
 *
 * Run: node generate_msp_xml.js
 */

'use strict';

const fs   = require('fs');
const path = require('path');

// ─── CSV helpers ──────────────────────────────────────────────────────────────

function parseCSVLine(line) {
  const result = [];
  let current  = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"')            { inQuotes = !inQuotes; }
    else if (ch === ',' && !inQuotes) { result.push(current.trim()); current = ''; }
    else                       { current += ch; }
  }
  result.push(current.trim());
  return result;
}

function parseDays(s) {
  const m = (s || '').match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

// ─── Date helpers ─────────────────────────────────────────────────────────────

function toXmlStart(s) {
  if (!s || !s.trim()) return '';
  const [mm, dd, yy] = s.trim().split('/');
  return `20${yy}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}T08:00:00`;
}

function toXmlFinish(s, isMilestone) {
  if (!s || !s.trim()) return '';
  const [mm, dd, yy] = s.trim().split('/');
  const t = isMilestone ? 'T08:00:00' : 'T17:00:00';
  return `20${yy}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}${t}`;
}

function daysToISO(days) {
  if (days === 0) return 'PT0H0M0S';
  return `PT${days * 8}H0M0S`;
}

// ─── XML escaping ─────────────────────────────────────────────────────────────

function esc(s) {
  return (s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;')
                  .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─── Parse CSV ────────────────────────────────────────────────────────────────

const csvPath = path.join(__dirname, 'Trinity_Project_Schedule.csv');
const raw     = fs.readFileSync(csvPath, 'utf8');
const lines   = raw.split(/\r?\n/).filter(l => l.trim());

const tasks = [];
for (let i = 1; i < lines.length; i++) {
  const row = parseCSVLine(lines[i]);
  if (!row[0] || isNaN(parseInt(row[0], 10))) continue;

  const id          = parseInt(row[0], 10);
  const outlineLevel = parseInt(row[1], 10);
  const name        = row[2] || '';
  const days        = parseDays(row[3]);
  const isMilestone = row[9] && row[9].trim().toLowerCase() === 'yes';

  const predecessors = (row[6] || '')
    .replace(/"/g,'')
    .split(',')
    .map(p => p.trim())
    .filter(p => p && !isNaN(parseInt(p, 10)))
    .map(p => parseInt(p, 10));

  tasks.push({
    id,
    outlineLevel,
    name,
    days,
    duration : daysToISO(days),
    start    : toXmlStart(row[4]),
    finish   : toXmlFinish(row[5], isMilestone || days === 0),
    predecessors,
    resources: row[7] || '',
    notes    : row[8] || '',
    milestone: isMilestone || days === 0,
  });
}

// ─── Identify summary tasks (next task has deeper outline level) ───────────────

const summarySet = new Set();
for (let i = 0; i < tasks.length - 1; i++) {
  if (tasks[i + 1].outlineLevel > tasks[i].outlineLevel) summarySet.add(tasks[i].id);
}

// ─── Build resource pool (split on ' + ' / ' +' / '+ ') ─────────────────────

const resourceSet = new Set();
tasks.forEach(t => {
  if (!t.resources) return;
  t.resources.split('+').map(r => r.trim()).filter(Boolean).forEach(r => resourceSet.add(r));
});
const resourceList = [...resourceSet].sort();
const resourceUID  = {};
resourceList.forEach((r, i) => { resourceUID[r] = i + 1; });

// ─── Generate XML ─────────────────────────────────────────────────────────────

const lines_xml = [];
const push = s => lines_xml.push(s);

push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`);
push(`<Project xmlns="http://schemas.microsoft.com/project">`);

// ── Project header ──
push(`  <SaveVersion>14</SaveVersion>`);
push(`  <Name>Trinity IT Carve-Out Project</Name>`);
push(`  <Title>Project Trinity</Title>`);
push(`  <Manager>IT PM</Manager>`);
push(`  <StartDate>2026-07-01T08:00:00</StartDate>`);
push(`  <FinishDate>2029-12-31T17:00:00</FinishDate>`);
push(`  <ScheduleFromStart>1</ScheduleFromStart>`);
push(`  <CalendarUID>1</CalendarUID>`);
push(`  <DefaultStartTime>08:00:00</DefaultStartTime>`);
push(`  <DefaultFinishTime>17:00:00</DefaultFinishTime>`);
push(`  <MinutesPerDay>480</MinutesPerDay>`);
push(`  <MinutesPerWeek>2400</MinutesPerWeek>`);
push(`  <DaysPerMonth>20</DaysPerMonth>`);
push(`  <DefaultTaskType>0</DefaultTaskType>`);
push(`  <DefaultFixedCostAccrual>3</DefaultFixedCostAccrual>`);
push(`  <CriticalSlackLimit>0</CriticalSlackLimit>`);
push(`  <CurrencySymbol>€</CurrencySymbol>`);
push(`  <CurrencyCode>EUR</CurrencyCode>`);
push(`  <CurrencyDigits>2</CurrencyDigits>`);

// ── Calendar ──
push(`  <Calendars>`);
push(`    <Calendar>`);
push(`      <UID>1</UID>`);
push(`      <Name>Standard</Name>`);
push(`      <IsBaseCalendar>1</IsBaseCalendar>`);
push(`      <IsBaselineCalendar>0</IsBaselineCalendar>`);
push(`      <WeekDays>`);
// Saturday (day 7) = non-working
push(`        <WeekDay>`);
push(`          <DayType>7</DayType>`);
push(`          <DayWorking>0</DayWorking>`);
push(`        </WeekDay>`);
// Sunday (day 1) = non-working
push(`        <WeekDay>`);
push(`          <DayType>1</DayType>`);
push(`          <DayWorking>0</DayWorking>`);
push(`        </WeekDay>`);
push(`      </WeekDays>`);
push(`    </Calendar>`);
push(`  </Calendars>`);

// ── Tasks ──
push(`  <Tasks>`);

// Task 0 — project summary (required by MS Project)
push(`    <Task>`);
push(`      <UID>0</UID>`);
push(`      <ID>0</ID>`);
push(`      <Name>Project Trinity</Name>`);
push(`      <Duration>PT27520H0M0S</Duration>`);
push(`      <DurationFormat>7</DurationFormat>`);
push(`      <Start>2026-07-01T08:00:00</Start>`);
push(`      <Finish>2029-12-31T17:00:00</Finish>`);
push(`      <Summary>1</Summary>`);
push(`      <Milestone>0</Milestone>`);
push(`      <OutlineLevel>0</OutlineLevel>`);
push(`      <CalendarUID>-1</CalendarUID>`);
push(`      <IgnoreResourceCalendar>0</IgnoreResourceCalendar>`);
push(`    </Task>`);

tasks.forEach(t => {
  const isSummary = summarySet.has(t.id);

  // Build combined notes (resources + notes)
  const noteParts = [];
  if (t.resources) noteParts.push(`Resources: ${t.resources}`);
  if (t.notes)     noteParts.push(t.notes);
  const noteText = noteParts.join(' | ');

  push(`    <Task>`);
  push(`      <UID>${t.id}</UID>`);
  push(`      <ID>${t.id}</ID>`);
  push(`      <Name>${esc(t.name)}</Name>`);
  push(`      <Duration>${t.duration}</Duration>`);
  push(`      <DurationFormat>7</DurationFormat>`);
  if (t.start)  push(`      <Start>${t.start}</Start>`);
  if (t.finish) push(`      <Finish>${t.finish}</Finish>`);
  push(`      <OutlineLevel>${t.outlineLevel}</OutlineLevel>`);
  push(`      <Summary>${isSummary ? 1 : 0}</Summary>`);
  push(`      <Milestone>${t.milestone ? 1 : 0}</Milestone>`);
  push(`      <CalendarUID>-1</CalendarUID>`);
  push(`      <IgnoreResourceCalendar>0</IgnoreResourceCalendar>`);
  push(`      <EffortDriven>0</EffortDriven>`);
  if (noteText) push(`      <Notes>${esc(noteText)}</Notes>`);

  t.predecessors.forEach(predId => {
    push(`      <PredecessorLink>`);
    push(`        <PredecessorUID>${predId}</PredecessorUID>`);
    push(`        <Type>1</Type>`);      // 1 = Finish-to-Start (default)
    push(`        <CrossProject>0</CrossProject>`);
    push(`        <LinkLag>0</LinkLag>`);
    push(`        <LagFormat>7</LagFormat>`);
    push(`      </PredecessorLink>`);
  });

  push(`    </Task>`);
});

push(`  </Tasks>`);

// ── Resources ──
push(`  <Resources>`);
resourceList.forEach((r, i) => {
  push(`    <Resource>`);
  push(`      <UID>${i + 1}</UID>`);
  push(`      <ID>${i + 1}</ID>`);
  push(`      <Name>${esc(r)}</Name>`);
  push(`      <Type>1</Type>`);           // 1 = Work resource
  push(`      <IsNull>0</IsNull>`);
  push(`      <CalendarUID>-1</CalendarUID>`);
  push(`      <IsEnterprise>0</IsEnterprise>`);
  push(`    </Resource>`);
});
push(`  </Resources>`);

// ── Assignments ──
push(`  <Assignments>`);
let aUID = 1;
tasks.forEach(t => {
  if (!t.resources) return;
  const parts = t.resources.split('+').map(r => r.trim()).filter(Boolean);
  parts.forEach(r => {
    const rid = resourceUID[r];
    if (!rid) return;
    const workHrs = t.days * 8;
    push(`    <Assignment>`);
    push(`      <UID>${aUID++}</UID>`);
    push(`      <TaskUID>${t.id}</TaskUID>`);
    push(`      <ResourceUID>${rid}</ResourceUID>`);
    push(`      <Units>1</Units>`);
    push(`      <Work>PT${workHrs}H0M0S</Work>`);
    if (t.start)  push(`      <Start>${t.start}</Start>`);
    if (t.finish) push(`      <Finish>${t.finish}</Finish>`);
    push(`    </Assignment>`);
  });
});
push(`  </Assignments>`);

push(`</Project>`);

// ─── Write file ───────────────────────────────────────────────────────────────

const xmlPath = path.join(__dirname, 'Trinity_Project_Schedule.xml');
fs.writeFileSync(xmlPath, lines_xml.join('\n'), 'utf8');

console.log(`✓ Written : ${xmlPath}`);
console.log(`  Tasks   : ${tasks.length}`);
console.log(`  Resources: ${resourceList.length}`);
console.log(`  Assignments: ${aUID - 1}`);
console.log(`  Milestones : ${tasks.filter(t => t.milestone).length}`);
console.log(`  Summary tasks: ${summarySet.size}`);
