const d = require('docx');
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak, LevelFormat, PageNumber, Footer
} = d;

const CW = 9360;

const P = (text, opts = {}) => new Paragraph({
  spacing: { after: opts.after === undefined ? 130 : opts.after, line: 270 },
  children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, size: 21, color: opts.color, font: 'Calibri' })],
});
const RP = (runs, opts = {}) => new Paragraph({
  spacing: { after: opts.after === undefined ? 130 : opts.after, line: 270 },
  children: runs.map(([t, o = {}]) => new TextRun({ text: t, bold: o.b, italics: o.i, color: o.c, size: 21, font: 'Calibri' })),
});
const H1 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 140 },
  children: [new TextRun({ text: t, bold: true, size: 28, color: '1A3A5C', font: 'Calibri' })],
});
const H2 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_2, spacing: { before: 210, after: 100 },
  children: [new TextRun({ text: t, bold: true, size: 23, color: '2C5A82', font: 'Calibri' })],
});
const BULR = (runs) => new Paragraph({
  numbering: { reference: 'bullets', level: 0 },
  spacing: { after: 80, line: 270 },
  children: runs.map(([t, o = {}]) => new TextRun({ text: t, bold: o.b, italics: o.i, color: o.c, size: 21, font: 'Calibri' })),
});
const BUL = (t) => BULR([[t]]);
const SP = (n = 100) => new Paragraph({ spacing: { after: n }, children: [new TextRun({ text: '', size: 2 })] });
const HR = () => new Paragraph({
  spacing: { before: 60, after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'BBBBBB' } },
  children: [new TextRun({ text: '', size: 2 })],
});

function mkTable(headers, rows, widths, opts = {}) {
  const cell = (text, { bold, shade, w, color } = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade, color: 'auto' } : undefined,
    margins: { top: 65, bottom: 65, left: 95, right: 95 },
    children: String(text).split('\n').map((line) => new Paragraph({
      spacing: { after: 0, line: 245 },
      children: [new TextRun({ text: line, bold, color, size: opts.size || 18, font: 'Calibri' })],
    })),
  });
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: '9AB2C6' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: '9AB2C6' },
      left: { style: BorderStyle.SINGLE, size: 4, color: '9AB2C6' },
      right: { style: BorderStyle.SINGLE, size: 4, color: '9AB2C6' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'C8D4DE' },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: 'C8D4DE' },
    },
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, { bold: true, shade: '1A3A5C', color: 'FFFFFF', w: widths[i] })) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, { w: widths[i], shade: ri % 2 ? 'F2F6F9' : undefined, bold: i === 0 && opts.boldFirst })) })),
    ],
  });
}

const CALLOUT = (title, body, colr) => {
  const c = colr || { bar: 'D9A441', fill: 'FDF6E7', txt: '8A6113' };
  return new Table({
    columnWidths: [CW], width: { size: CW, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: c.bar },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: c.bar },
      left: { style: BorderStyle.SINGLE, size: 18, color: c.bar },
      right: { style: BorderStyle.SINGLE, size: 2, color: c.bar },
      insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: c.fill, color: 'auto' },
        margins: { top: 120, bottom: 120, left: 150, right: 150 },
        children: [
          new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: title, bold: true, size: 21, color: c.txt, font: 'Calibri' })] }),
          ...body.map((b) => new Paragraph({ spacing: { after: 60, line: 265 }, children: [new TextRun({ text: b, size: 20, font: 'Calibri' })] })),
        ],
      })],
    })],
  });
};

const ch = [];

// ---------------------------------------------------------------- TITLE
ch.push(new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: 'PRODUCT REQUIREMENTS DOCUMENT', bold: true, size: 18, color: '7A8C9A', font: 'Calibri' })] }));
ch.push(new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: 'Portable Aimbot Target', bold: true, size: 50, color: '1A3A5C', font: 'Calibri' })] }));
ch.push(new Paragraph({ spacing: { after: 180 }, children: [new TextRun({ text: 'Shared team training equipment — switch it on and practise. Packs into one checked bag.', size: 23, color: '4A5A68', italics: true, font: 'Calibri' })] }));

ch.push(mkTable(
  ['Field', 'Value'],
  [
    ['Primary users', 'Aimbot team members — the people testing and tuning auto-aim'],
    ['Builder / maintainer', 'Solo (mechanical, EE, firmware, software)'],
    ['Version / date', '2.0 — 28 July 2026'],
    ['Primary deadline', 'Workshop-usable by end of November 2026 (~17 weeks)'],
    ['Secondary deadline', 'Travel-ready kit by end of February 2027, ahead of the RMUL 2027 cycle (Q1)'],
    ['Transport constraint', 'One standard checked bag: 158 cm linear, ≤ 23 kg, case included'],
    ['Budget', '~$910 all phases, or $550 for Phases 1–2 if staged (§10)'],
    ['Status', 'Pending decisions D1–D6 and open questions Q1–Q7 (§12)'],
  ],
  [2300, 7060], { boldFirst: true }
));
ch.push(SP(200));

// ---------------------------------------------------------------- 1 SUMMARY
ch.push(H1('1. Summary'));

ch.push(P('Auto-aim can only be tested against a moving target. Producing one today requires a second complete robot plus someone to drive it, so testing is gated on hardware, people and battery charge — and consequently happens rarely, in short windows, under whatever conditions exist. Regressions get found at competitions instead of in the workshop.'));

ch.push(P('This project builds a motorised spinning armor target that removes that dependency. It also holds a commanded speed, so two test runs a week apart are comparable, which a human-driven robot never is. And it packs into one checked bag so the same known target can be set up in a competition pit — where unfamiliar lighting and a compressed schedule make a five-minute sanity check worth more than anywhere else.'));

ch.push(RP([['This is shared team equipment, not a personal rig. ', { b: true }], ['Any aimbot team member should be able to take it out, use it, and put it away without asking the person who built it. That constraint shapes the mechanical design, the UI and the documentation more than any measurement requirement does.']]));

ch.push(CALLOUT('Recommendation', [
  'Build the spinning head on a heavy fixed workshop base first and get it into team use by Week 6. That single deliverable removes the robot dependency and is most of the practical value. Add the workshop rail second, instrumented plates and the UI third, and the travel stand and case fourth.',
  'Design the head against the transport envelope from day one even though the case is built last. Portability cannot be retrofitted; every part is either small enough on the day it is designed or it never will be.',
]));

// ---------------------------------------------------------------- 2 USERS
ch.push(H1('2. Users'));

ch.push(mkTable(
  ['User', 'Job to be done', 'Success looks like'],
  [
    ['Aimbot team member,\nworkshop (primary)', 'Test and tune auto-aim against a moving target, alone, on a weeknight.', 'Sets it up and starts testing without asking anyone how it works.'],
    ['Aimbot team member,\ncompetition (primary)', 'Confirm tracking still behaves under venue lighting before a match.', 'Deployed and running in under 10 minutes on unfamiliar ground.'],
    ['Builder / maintainer\n(secondary)', 'Keep it running and fix it when it breaks.', 'Spends time improving it, not explaining it.'],
    ['Next season\'s team\n(tertiary)', 'Inherit working equipment rather than a half-finished project.', 'Still works after a year in storage; spares are documented.'],
  ],
  [2200, 3600, 3560]
));
ch.push(SP(140));

ch.push(RP([['The primary user did not build this and will not read this document. ', { b: true }], ['They will pick it up at 9 p.m. with a robot half-assembled behind them. Anything that requires knowing a trick, remembering an order, or asking a question is a defect — not a documentation gap. This is why §6.3 spends as much effort on unambiguous assembly as on structural stiffness, and why the boot self-test in §8 exists.']]));

// ---------------------------------------------------------------- 3 PRINCIPLES
ch.push(H1('3. Principles and non-goals'));

ch.push(H2('3.1 Principles'));
ch.push(BULR([['No tribal knowledge. ', { b: true }], ['If using it correctly depends on something only the builder knows, the design is wrong. Keying, labelling and self-tests replace explanation.']]));
ch.push(BULR([['Dependability beats capability. ', { b: true }], ['A target that spins reliably and does nothing else beats one that measures everything and works two sessions in three.']]));
ch.push(BULR([['Slot to locate, latch to clamp. ', { b: true }], ['Slots position parts; they do not hold them. Every joint separates the two functions (§6.3).']]));
ch.push(BULR([['Nothing loose. ', { b: true }], ['Every surviving fastener is captive. A dropped screw on a competition floor should not be able to end the trip.']]));
ch.push(BULR([['The case is the checklist. ', { b: true }], ['Foam cut to each part means an empty pocket is a visible missing part.']]));

ch.push(H2('3.2 Non-goals'));
ch.push(BUL('Not a competition-legal robot — it never enters the field of play.'));
ch.push(BUL('Not a general-purpose test range. It spins plates and translates on a rail; it does not simulate a driving robot or terrain.'));
ch.push(BUL('Not a referee-system replacement. Where they disagree in the workshop, the referee system is authoritative for hit counting.'));
ch.push(BUL('No auto-tuning in this scope — noted as an unscheduled future phase in §9.'));
ch.push(BUL('No mobile app, no cloud, no accounts. The UI is a local web page.'));

// pagebreak removed

// ---------------------------------------------------------------- 4 ARCHITECTURE
ch.push(H1('4. Architecture: one head, three bases'));

ch.push(P('Treat the spinning head as the product and everything below it as an interchangeable base sharing one mechanical and electrical interface.'));

ch.push(mkTable(
  ['Module', 'What it is', 'Constraints', 'Phase'],
  [
    ['Head (the product)', 'Motor, hub, four plate carriers, plates, rotating-frame electronics.', 'Full transport envelope and weight budget. Designed once, changed rarely.', '1'],
    ['Workshop base', 'Heavy ballasted pedestal. Deliberately crude.', 'None. Weight is a feature here, not a cost.', '1'],
    ['Rail base', 'Long rail with a driven carriage the head mounts to.', 'Workshop only. Never flies.', '2'],
    ['Travel base', 'Collapsing stand, packs flat.', 'The whole transport envelope. Hardest mechanical problem here.', '4'],
  ],
  [1600, 2900, 3360, 1500]
));
ch.push(SP(140));

ch.push(P('Three things follow. A usable machine reaches the team in about six weeks, because the first base is a lump of steel needing no design effort. The travel base is designed against a head whose geometry and mass have stopped changing, which is when a packing problem becomes tractable. And if the travel base proves too hard, the workshop keeps its equipment — the failure is contained.'));

ch.push(RP([['One specification carries the scheme: ', { b: true }], ['a defined bolt or slot pattern, a vertical datum, and one connector carrying 24 V and data. Freeze it in Phase 1. Every hour spent on it in week 3 saves a week in month 5.']]));

// ---------------------------------------------------------------- 5 METRICS
ch.push(H1('5. Goals and success metrics'));

ch.push(mkTable(
  ['#', 'Metric', 'Target', 'Why this measure'],
  [
    ['M1', 'Setup time, case to first shot', '≤ 10 min, one person, no tools', 'Decides whether it actually gets used'],
    ['M2', 'Use by someone who did not build it', 'A team member sets up and runs a full session unaided, first attempt', 'The primary user test. Tests the design, not the builder\'s memory'],
    ['M3', 'Sessions between tool-requiring failures', '≥ 20', 'Encodes "switch it on and start using it"'],
    ['M4', 'Packed size', '≤ 158 cm linear, case included', 'Binary pass/fail at the check-in desk'],
    ['M5', 'Packed weight', '≤ 23 kg, target ≤ 20 kg', '3 kg margin absorbs spares and a heavier case than planned'],
    ['M6', 'Spin speed consistency', '≤ 2% from commanded, 0.25–3.0 rev/s', 'Two sessions a week apart must be comparable'],
    ['M7', 'Hit localisation accuracy', '≤ 25 mm RMS, ≥ 95% correct quadrant', 'Honest limit for piezo TDOA on a small plate (§8.2)'],
    ['M8', 'Adoption', '≥ 8 team practice sessions in the first month after Phase 1, by at least two different people', 'A tool one person uses is a rig; a tool the team uses is equipment'],
  ],
  [450, 2450, 2700, 3760]
));
ch.push(SP(140));
ch.push(RP([['M2 and M8 are the pair that matter. ', { b: true }], ['Every other metric can be satisfied by a machine that only its builder can operate. If teammates are not using it within a month of Phase 1, something is wrong that no specification captured — stop and find out what.']]));

// pagebreak removed

// ---------------------------------------------------------------- 6 MECHANICAL
ch.push(H1('6. Transport and mechanical requirements'));

ch.push(H2('6.1 The transport envelope'));
ch.push(P('Checked baggage on carriers flying Singapore to China is limited to 158 cm of linear dimension — length plus width plus height summed — and 23 kg in economy. The critical detail is that 158 cm is a sum, not a length, so case shape is a free design variable.'));

ch.push(mkTable(
  ['Case shape (cm)', 'Longest internal member', 'Comment'],
  [
    ['70 × 48 × 40', '~66 cm', 'Conventional suitcase. Easy to handle, poor length.'],
    ['85 × 45 × 28', '~81 cm', 'Balanced. Current working assumption.'],
    ['105 × 32 × 20', '~101 cm', 'Ski-case proportions. Best length, awkward, more bending risk.'],
  ],
  [2100, 2400, 4860]
));
ch.push(SP(140));

ch.push(RP([['Choose the shape after the largest irreducible part is known, ', { b: true }], ['not before. Design the head, find the part you cannot shrink further, then pick the proportions that fit it with least wasted volume.']]));

ch.push(CALLOUT('Build to the small armor format (~135 mm), not the large (~235 mm)', [
  'The plates are custom, so the format is yours to choose, and small wins three ways. Four small plates occupy well under half the volume of four large ones. It is the harder and more realistic target. And TDOA still works comfortably — a flexural wave crosses a 135 mm plate in roughly 70 µs, which an STM32 timer resolves with enormous margin (§8.2).',
]));

ch.push(H2('6.2 Weight is the binding constraint'));
ch.push(P('Size fails visibly — a part fits the case or it does not. Weight accumulates invisibly and is discovered at check-in. Preliminary budget, flying items only:'));

ch.push(mkTable(
  ['Subsystem', 'kg', 'Subsystem', 'kg'],
  [
    ['GM6020 motor (confirmed)', '0.47', 'Travel stand', '3.50'],
    ['Head hub, arms, carriers', '1.80', 'Spares kit', '0.80'],
    ['4 × plates with sensors', '1.60', 'Case and foam', '4.50'],
    ['Rotating electronics + slip ring', '0.30', 'Cables and PSU lead', '0.40'],
    ['Base electronics and PSU', '1.50', 'Contingency at 25%', '3.72'],
    ['', '', 'PROJECTED TOTAL', '18.59'],
  ],
  [3100, 1200, 3860, 1200], { boldFirst: true }
));
ch.push(SP(140));
ch.push(P('This passes against 23 kg but not comfortably. The stand and case are 44% of flying mass and the two least-defined items, so build and weigh both early in Phase 4 rather than at the end. Track weight continuously against this budget.'));

ch.push(CALLOUT('Hard rule: no lithium batteries in checked baggage', [
  'Loose lithium packs are prohibited in checked baggage worldwide, and carry-on is capped around 100 Wh. A standard RM battery exceeds this comfortably. Nothing battery-powered flies with this kit under any packing arrangement.',
  'Consequence: a universal-input mains PSU (100–240 V covers both Singapore and mainland China) as the primary supply, with a DC input that also accepts a borrowed RM pack on site. Venue power becomes an open risk, not an assumption — Q3.',
], { bar: 'C0504D', fill: 'FDEDEC', txt: '922B21' }));

ch.push(H2('6.3 Requirements'));
ch.push(BULR([['REQ-M1 ', { b: true }], ['Four plate carriers at 90° on a hub driven directly by a GM6020. Direct drive removes gearbox backlash, and the motor\'s absolute encoder gives plate angle for free at 8192 counts per revolution.']]));
ch.push(BULR([['REQ-M2 ', { b: true }], ['0.25–3.0 rev/s held within 2%, plus a static hold mode. The GM6020 gives 1.2 N·m and 320 rpm unloaded, so top speed has margin; spin-up is the tighter case. At a plausible head inertia near 0.08 kg·m², reaching 3 rev/s takes about 1.5 s. Treat head inertia as a budgeted quantity, not an outcome.']]));
ch.push(BULR([['REQ-M3 ', { b: true }], ['Plate carriers detach from arms, and arms from the hub, without tools. These are the largest parts and must nest flat in the case.']]));
ch.push(BULR([['REQ-M4 ', { b: true }], ['Route power and data through the motor\'s 18 mm hollow shaft — the wiring path through the rotation axis already exists, so no cable management around a spinning assembly is needed.']]));
ch.push(BULR([['REQ-M5 ', { b: true }], ['Balance within a defined residual imbalance so the machine does not walk or resonate at 3 rev/s. Provide balance mass positions, and document which reduced-plate configurations are safe to run.']]));
ch.push(BULR([['REQ-M6 ', { b: true }], ['One frozen head-to-base interface. All three bases implement it.']]));
ch.push(BULR([['REQ-M7 ', { b: true }], ['Rail base: belt drive with a stationary motor at one end. A moving motor needs a moving power feed; a lead screw over several metres is expensive and slow.']]));
ch.push(BULR([['REQ-M8 ', { b: true }], ['Travel base: assembles without tools, packs flat, and stable enough that a 3 rev/s head does not walk it across a smooth venue floor. Stability from footprint and from ballast added on site — a water bottle strapped to the base weighs nothing in the suitcase.']]));

ch.push(H2('6.4 Joints: the central mechanical tension'));
ch.push(P('Slotting is right for assembly speed and wrong for a machine that vibrates. A slotted joint locates parts precisely and quickly, but does not by itself generate the preload that stops a joint fretting loose under cyclic load — and a spinning head taking repeated projectile impacts is exactly that load. A design relying on slot friction alone works perfectly on the bench and is loose by the third session.'));
ch.push(BULR([['REQ-M9 ', { b: true }], ['Every structural joint separates locating from clamping. Geometry — tabs, slots, dowels, tapers — sets position. A dedicated element — draw latch, cam lock, quarter-turn fastener, or a self-tightening wedge — provides preload. Neither does the other\'s job.']]));
ch.push(BULR([['REQ-M10 ', { b: true }], ['Every remaining fastener is captive in its part.']]));
ch.push(BULR([['REQ-M11 ', { b: true }], ['Every joint is verifiable by sight or feel. A visibly open latch beats a torque spec nobody will check — especially for a user who did not build the machine.']]));
ch.push(BULR([['REQ-M12 ', { b: true }], ['Assembly is unambiguous. Parts that could be fitted the wrong way round are keyed so they cannot be, or made genuinely symmetric so it does not matter. Every part is labelled to match its foam pocket and the setup card. This is what makes M2 achievable.']]));

ch.push(H2('6.5 Durability and the case'));
ch.push(BULR([['Plate faces ', { b: true }], ['absorb thousands of 17 mm impacts at up to 30 m/s. Make the impact face a separate, cheap, replaceable layer rather than part of the sensor assembly.']]));
ch.push(BULR([['Piezo bonds ', { b: true }], ['are the insidious problem: repeated impact fatigues an adhesive bond, and a partially debonded sensor does not fail cleanly — it shifts its response and quietly corrupts localisation while still appearing to work. Sensors are field-replaceable modules, and degradation is detected actively (§8.2), not noticed eventually.']]));
ch.push(BULR([['REQ-M13 ', { b: true }], ['Custom foam, one cut pocket per labelled part. Assembly order printed inside the lid — instructions attached to the case cannot be left at home.']]));
ch.push(BULR([['REQ-M14 ', { b: true }], ['Survives a drop test representative of baggage handling. Test before the trip, not at the destination. The spares kit travels inside the case as a fixed part of the packout.']]));

// pagebreak removed

// ---------------------------------------------------------------- 7/8 ELECTRONICS
ch.push(H1('7. Electronics and firmware'));

ch.push(H2('7.1 The rotating frame'));
ch.push(P('Four plates, each with four piezo sensors, all rotating. Sixteen analogue channels carrying microsecond-scale timing must become useful data somewhere, and the spinning interface is where that naturally goes wrong.'));
ch.push(RP([['REQ-E1 — do the timing work on the rotating side. ', { b: true }], ['Analogue front end and capture MCU sit on the head. Only digested impact events — plate ID, coordinates, timestamp, energy — cross the interface, as small digital packets.']]));
ch.push(P('This removes a class of failure. Slip-ring contact noise is fatal to microsecond timing on an analogue signal and irrelevant to a digital packet or a DC rail. It also cuts conductors crossing the interface from sixteen-plus to about four, which makes the slip ring smaller, cheaper and longer-lived — and a slip ring is a wear item, so fewer contacts is directly fewer things that degrade.'));
ch.push(BULR([['REQ-E2 ', { b: true }], ['Slip ring carries 24 V and a digital link only, passing through the hollow shaft on the rotation axis.']]));
ch.push(BULR([['REQ-E3 ', { b: true }], ['A wireless data link is an acceptable alternative that eliminates a wear item — but for data only; power still crosses mechanically. Decide once measured slip-ring reliability exists (D4).']]));

ch.push(H2('7.2 Hit localisation'));
ch.push(P('Four piezo sensors near the corners of each plate; an impact launches a flexural wave; differences in arrival time locate the source by multilateration.'));
ch.push(BULR([['The electronics are comfortable. ', { b: true }], ['Flexural waves travel at roughly 1000–3000 m/s, crossing a 135 mm plate in around 45–135 µs. Four comparator outputs into an STM32 timer\'s input-capture channels resolve arrival times to nanoseconds. The clock is nowhere near the limitation.']]));
ch.push(BULR([['The limitation is physics. ', { b: true }], ['These waves disperse and reverberate off plate edges, so the threshold-crossing instant shifts with impact energy and location. Published work reports errors of a few percent of plate dimension under controlled conditions — hence a 25 mm RMS target, not a sub-centimetre one. Any design promising millimetres will fail its own acceptance test.']]));
ch.push(BULR([['REQ-E4 ', { b: true }], ['Calibrate empirically against a known grid of tap positions; store the map with the plate. An empirical map beats an analytical wave-speed model, and each plate needs its own.']]));
ch.push(BULR([['REQ-E5 ', { b: true }], ['Bench go/no-go before integration: build one instrumented plate and characterise it. If calibrated RMS exceeds 25 mm, stop — do not build four of something that does not work.']]));
ch.push(BULR([['REQ-E6 ', { b: true }], ['Detect your own sensor degradation. Monitor per-channel response consistency and flag any channel drifting from its calibrated baseline. This turns silent data corruption into a visible warning — essential when the user did not build the machine and has no basis for suspecting the data.']]));

ch.push(H2('7.3 Referee system and time base'));
ch.push(BULR([['REQ-E7 ', { b: true }], ['Use the referee system in the workshop as the reference validating custom plate hit detection. Two independent counts that agree is strong evidence both work.']]));
ch.push(BULR([['REQ-E8 ', { b: true }], ['Do not make the travel configuration depend on it. Custom plates must detect hits standalone; a referee set at the venue is a bonus. Referee equipment is generally issued per registered robot and its use on a non-robot fixture may not be permitted — confirm early (Q4).']]));
ch.push(BULR([['REQ-E9 ', { b: true }], ['Attributing an impact to a plate angle within 5° at 3 rev/s needs hit and encoder pairing to about 2 ms, which ordinary software timestamps meet with room to spare. Leave a pin, pad and connector position so a hardware trigger could be added later — auto-tuning would tighten this to roughly ±200 µs, and reserving the path now costs an hour where retrofitting costs a board revision.']]));

// ---------------------------------------------------------------- 8 UI
ch.push(H1('8. User interface'));

ch.push(P('A local web page served by the base controller, opened in a browser over USB or the local network. No installation and no dependencies — which matters because the primary user is on a borrowed laptop in a pit area where installing software is impossible.'));

ch.push(mkTable(
  ['Screen', 'Shows', 'Priority'],
  [
    ['Live', 'Spin rate vs. commanded, speed control, start/stop/E-stop, four plate outlines with impacts appearing live.', 'Must'],
    ['Health', 'Boot self-test: all 16 sensor channels, motor comms, encoder, referee link. Green or red, per item.', 'Must'],
    ['Session', 'Hit count and distribution per plate, group centre and spread, hit rate over the session.', 'Must'],
    ['History', 'Past sessions, comparison, CSV export.', 'Should'],
    ['Calibration', 'Guided per-plate tap-grid calibration and the resulting accuracy figure.', 'Should'],
  ],
  [1300, 6560, 1500]
));
ch.push(SP(140));

ch.push(RP([['The Health screen is what makes the machine trustworthy to someone who did not build it. ', { b: true }], ['Without it, a debonded sensor produces plausible but wrong data and the user has no way to know which sessions to discard. A power-up check reporting every channel good or bad converts an ambiguous machine into a dependable one. Build it in Phase 3 alongside the plates, not later.']]));

// pagebreak removed

// ---------------------------------------------------------------- 9 PHASES
ch.push(H1('9. Phases and gates'));

ch.push(mkTable(
  ['Phase', 'Weeks', 'Scope', 'Exit criterion'],
  [
    ['1 — Workshop target', '1–6', 'Head, workshop base, motor control, speed hold, safety, stock plates. Frozen interface. Every head part checked against the transport envelope.', 'A team member other than the builder runs a full practice session with it, unaided.'],
    ['2 — Rail', '5–10', 'Long workshop rail, driven carriage, unified translation and spin control.', 'Head translates and spins together, reliably, across a full session.'],
    ['3 — Plates + UI', '8–17', 'One prototype plate and bench characterisation first, then four. Rotating-frame electronics, slip ring, web UI, boot self-test, referee cross-check.', '≤ 25 mm RMS on all four plates; a teammate uses the UI without explanation.'],
    ['4 — Travel kit', '14–26', 'Collapsing stand, custom case and foam, part labelling, packout, drop test, timed dry run.', 'Under 23 kg and 158 cm; assembled from case in under 10 min by someone who did not build it.'],
    ['5 — Auto-tuning', 'unscheduled', 'Synchronised logging, hardware time base, offline parameter replay. Needs the ±200 µs work deferred in §7.3.', 'Not scheduled. Pick up only when Phases 1–4 are genuinely finished.'],
  ],
  [1750, 1000, 3450, 3160]
));
ch.push(SP(140));

ch.push(P('Phases 1 and 2 overlap deliberately — the rail is largely independent once the interface is frozen, and it fills the gaps while waiting for parts. Phase 3 starts with a single plate for one reason: to find out whether the approach works before committing to four.'));

ch.push(mkTable(
  ['Gate', 'Week', 'Test', 'Action on failure'],
  [
    ['Adoption', '6', 'Does a team member reach for it without being prompted?', 'Understand why not before adding a rail, sensors or a suitcase to it'],
    ['Plate', '10', 'Prototype plate ≤ 25 mm RMS on held-out grid positions', 'Do not build four. Ship with stock plates and referee counting — still good training equipment'],
    ['Packing', '20–21', 'Full dry-run packout under 23 kg and 158 cm', 'Decide deliberately: oversize fee, freight, or workshop-only. Not at the check-in desk'],
    ['Dry run', '24', 'Teammate assembles from the case unaided, under 10 min', 'The design is ambiguous. Fix keying and labelling, re-run'],
    ['Reliability', 'any', 'Two consecutive sessions needing tools to get running', 'Stop adding features. Everything here is downstream of it working'],
  ],
  [1200, 800, 3400, 3960]
));

// ---------------------------------------------------------------- 10 BUDGET
ch.push(H1('10. Budget'));

ch.push(mkTable(
  ['Group', 'Phase', 'Est.', 'Group', 'Phase', 'Est.'],
  [
    ['Head', '1', '$210', 'Rotating electronics', '3', '$80'],
    ['Workshop base', '1', '$40', 'Travel stand', '4', '$100'],
    ['Base electronics', '1', '$110', 'Case and foam', '4', '$90'],
    ['Rail', '2', '$190', 'Spares kit', '4', '$50'],
    ['Instrumented plates', '3', '$150', 'TOTAL', '', '$1,020'],
    ['', '', '', 'Less GM6020 if a spare exists', '', '−$110'],
    ['', '', '', 'LIKELY SPEND', '', '$910'],
  ],
  [2000, 800, 900, 3200, 800, 1660], { boldFirst: true }
));
ch.push(SP(140));

ch.push(RP([['Staging option. ', { b: true }], ['Phases 1 and 2 together are $550 and deliver the training tool that removes the robot dependency — most of the practical value. Phases 3 and 4 add $470 for measurement and portability. Funding the first half now and deciding the second at the Week 8 review also defers the plate spend until the Week 10 gate has said whether TDOA works.']]));

// ---------------------------------------------------------------- 11 RISKS
ch.push(H1('11. Risks'));

ch.push(mkTable(
  ['#', 'Risk', 'Sev.', 'Mitigation'],
  [
    ['R1', 'Slotted joints loosen under vibration and impact', 'Critical', 'REQ-M9: separate locating from clamping. Endurance run and full joint re-check in Phase 1, not Phase 4.'],
    ['R2', 'Only the builder can operate or fix it', 'Critical', 'M2 makes this a tested requirement, not an aspiration. Keying, labelling, self-test and a setup card in the case. Have a teammate run Phase 1 acceptance.'],
    ['R3', 'Weight creep pushes the packed kit past 23 kg', 'Critical', 'Weigh every part as it is made. Stand and case are 44% of flying mass and least defined — build and weigh them early in Phase 4.'],
    ['R4', 'Piezo bonds fatigue and silently corrupt localisation', 'Critical', 'REQ-E6 per-channel drift detection, field-replaceable modules, boot self-test on the Health screen.'],
    ['R5', 'Solo builder, in-season workload, project stalls part-built', 'Critical', 'Phase 1 delivers standalone value in 6 weeks; every later phase is independently useful. Stopping at any phase is acceptable.'],
    ['R6', 'Travel stand cannot be both packable and stable', 'High', 'Contained by the architecture — workshop use is unaffected. Ballast filled on site rather than stiffness in the suitcase.'],
    ['R7', 'No usable mains power at the venue', 'High', 'Dual DC input accepting a borrowed pack. No battery can be flown regardless. Confirm venue power (Q3).'],
    ['R8', 'Case fails baggage handling; parts arrive damaged', 'High', 'Drop test before travelling. Spares inside the case. Irreplaceable electronics in hand luggage.'],
    ['R9', 'GM6020 torque marginal once the head is fully built', 'High', 'Budget head inertia before building. Keep plate radius small — inertia scales with r², the cheapest lever available.'],
    ['R10', 'Injury from the spinning head or projectiles', 'High', 'Hardware E-stop cutting motor power directly. Balance verification before first full-speed run. Captive fasteners. Eye protection.'],
    ['R11', 'Slip ring wears or becomes intermittent', 'High', 'REQ-E1 keeps only power and digital across the interface. REQ-E3 keeps a wireless fallback open.'],
    ['R12', 'Referee gear unusable on a non-robot fixture at a venue', 'Medium', 'Custom plates work standalone, so referee is a workshop cross-check, not a dependency (Q4).'],
  ],
  [450, 3200, 900, 4810]
));
ch.push(SP(140));

ch.push(RP([['R2 is new to this version and it is structural. ', { b: true }], ['A solo builder producing equipment that a team depends on creates a machine only one person can operate, and that person will be busy during a competition. The mitigation is not documentation written at the end — it is designing so that documentation is barely needed, and then testing that claim with someone who was not involved.']]));

// ---------------------------------------------------------------- 12 DECISIONS
ch.push(H1('12. Decisions and open questions'));

ch.push(H2('12.1 Decisions'));
ch.push(mkTable(
  ['ID', 'Decision', 'Rationale'],
  [
    ['D1', 'Head is the product; three bases share one frozen interface', 'Usable machine in 6 weeks; travel base designed against settled geometry; contained failure'],
    ['D2', 'Build to the small armor format (~135 mm)', 'Packs far smaller, harder and more realistic target, TDOA still comfortable'],
    ['D3', 'TDOA capture and processing on the rotating head; only digital events cross the interface', 'Removes slip-ring noise from the timing path; conductors drop from 16+ to ~4'],
    ['D4', 'Slip ring for power and data initially; wireless data reserved as fallback', 'Fewer unknowns in Phase 3; revisit once measured reliability exists'],
    ['D5', 'Referee system is a workshop cross-check, never a travel dependency', 'Weight, team-property and battery constraints all argue against flying it'],
    ['D6', 'Time base built to ±2 ms; upgrade path to ±200 µs reserved, not built', 'Impact-angle attribution needs only 2 ms; reserving the path is nearly free'],
  ],
  [450, 3400, 5510]
));
ch.push(SP(140));

ch.push(H2('12.2 Open questions'));
ch.push(mkTable(
  ['ID', 'Question', 'Needed by'],
  [
    ['Q1', 'Which competition and date is the travel deadline? RMUL 2026 mainland sites ran March–May 2026, so this assumes the 2027 cycle around March 2027. If the real date is earlier, Phase 4 must move ahead of Phase 3.', 'Week 2'],
    ['Q7', 'Is there 3D printing and machining access, or is everything cut by hand?', 'Week 1'],
    ['Q6', 'What plate height above the floor matches the robots most worth practising against?', 'Week 3'],
    ['Q2', 'Is $910 approvable, or staged — $550 now for Phases 1–2, the rest decided at Week 8?', 'Week 4'],
    ['Q5', 'How long is the workshop rail, and what mounting does the space allow?', 'Week 5'],
    ['Q3', 'Is mains power available near a usable test area in a competition pit?', 'Week 8'],
    ['Q4', 'Can referee gear be used on a non-robot fixture at a venue, and is a spare set available?', 'Week 8'],
  ],
  [450, 7450, 1460]
));

ch.push(HR());
ch.push(P('Companion artefact: the phased roadmap workbook — week-by-week plan, weight budget tracker, case sizing and packing manifest, BOM, risk register and open-questions tracker.', { italics: true, color: '666666' }));

// -------------------------------------------------------------- BUILD
const doc = new Document({
  creator: 'Product Management',
  title: 'Portable Aimbot Target — PRD',
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 380, hanging: 190 } } } },
      ],
    }],
  },
  styles: { default: { document: { run: { font: 'Calibri', size: 21 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1150, bottom: 1150, left: 1440, right: 1440 } } },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: 'Portable Aimbot Target PRD  |  ', size: 16, color: '999999' }),
                     new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '999999' })],
        })],
      }),
    },
    children: ch,
  }],
});

Packer.toBuffer(doc).then((b) => {
  fs.writeFileSync('/sessions/great-vigilant-edison/mnt/outputs/Portable-Aimbot-Target-PRD.docx', b);
  console.log('written');
});
