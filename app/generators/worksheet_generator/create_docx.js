const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, LevelFormat, BorderStyle
} = require('docx');
const fs = require('fs');

// ── Helpers ──────────────────────────────────────────────────────────────────

const bold = (text) => new TextRun({ text, bold: true });
const normal = (text) => new TextRun({ text });

const p = (...children) => new Paragraph({ children });
const meta = (label, value) => new Paragraph({
  children: [bold(label + ': '), normal(value)]
});
const heading = (text, level = HeadingLevel.HEADING_2) =>
  new Paragraph({ text, heading: level });
const plain = (text) => new Paragraph({ children: [normal(text)] });
const qPara = (qNum, text) => new Paragraph({
  children: [bold(`Q${qNum}. `), normal(text)]
});
const option = (letter, text) => new Paragraph({
  children: [normal(`${letter}) ${text}`)]
});
const assertion = (text) => new Paragraph({
  children: [bold('Assertion: '), normal(text)]
});
const reason = (text) => new Paragraph({
  children: [bold('Reason: '), normal(text)]
});
const answer = (text) => new Paragraph({
  children: [bold('Answer: '), normal(text)]
});
const answerLabel = () => new Paragraph({ children: [bold('Answer:')] });
const answerLine = (text) => new Paragraph({
  children: [normal(text)],
  indent: { left: 360 }
});
const blank = () => new Paragraph({ text: '' });

// ── Document ─────────────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Calibri', size: 22 } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal',
        run: { size: 28, bold: true, font: 'Calibri' },
        paragraph: { spacing: { before: 200, after: 100 } }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal',
        run: { size: 24, bold: true, font: 'Calibri', color: '0070C0' },
        paragraph: { spacing: { before: 200, after: 80 } }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 720, right: 720, bottom: 720, left: 720 }
      }
    },
    children: [

      // ── Metadata ────────────────────────────────────────────────────────────
      new Paragraph({
        text: 'Worksheet Input Document',
        heading: HeadingLevel.HEADING_1
      }),
      blank(),

      meta('Chapter', 'Chapter 1 – Wonderful World of Science'),
      meta('Class', '6'),
      meta('Subject', 'Science'),
      meta('Worksheet', 'Worksheet 3 Key (Everyday Scientists & Collaboration)'),
      blank(),

      // ── Section A: MCQ ──────────────────────────────────────────────────────
      new Paragraph({ text: 'A. MULTIPLE CHOICE QUESTIONS (MCQ)', heading: HeadingLevel.HEADING_2 }),

      qPara(1, 'Why does a bicycle repair person use a bowl of water to check a flat tyre?'),
      option('a', 'To clean the tyre.'),
      option('b', 'To test their guess about where the air is leaking.'),
      option('c', 'To observe if the rubber changes color.'),
      option('d', 'Because water makes the tyre stronger.'),
      answer('b) To test their guess about where the air is leaking'),
      blank(),

      qPara(2, 'What is the "first and foremost" thing needed to learn science well?'),
      option('a', 'A laboratory'),
      option('b', 'Expensive books'),
      option('c', 'To be curious and observe surroundings keenly'),
      option('d', 'To memorize all facts.'),
      answer('c) To be curious and observe surroundings keenly'),
      blank(),

      // ── Section B: FIB ──────────────────────────────────────────────────────
      new Paragraph({ text: 'B. FILL IN THE BLANKS (FIB)', heading: HeadingLevel.HEADING_2 }),

      qPara(3, 'Anyone who follows the __________ method to solve problems or discover new things is working like a scientist.'),
      answer('scientific'),
      blank(),

      qPara(4, 'To be a wise person, the text suggests you must be a "__________" person.'),
      answer('whys'),
      blank(),

      // ── Section C: T/F ──────────────────────────────────────────────────────
      new Paragraph({ text: 'C. TRUE OR FALSE (T/F)', heading: HeadingLevel.HEADING_2 }),

      qPara(5, 'An electrician trying to find out why a bulb is not working is acting like a scientist.'),
      answer('True – They use logic to find whether the problem is the bulb or the switch.'),
      blank(),

      qPara(6, 'Scientific discovery is an ending jigsaw puzzle with a limited number of pieces.'),
      answer('False – The puzzle is "unending" and there is no limit to discovery.'),
      blank(),

      // ── Section D: Assertion–Reason ─────────────────────────────────────────
      new Paragraph({ text: 'D. ASSERTION–REASON', heading: HeadingLevel.HEADING_2 }),

      qPara(7, 'Assertion-Reason Question:'),
      assertion('Collaboration is important in science.'),
      reason('It is more fun and effective to discover things together with friends.'),
      option('a', 'Both Assertion and Reason are true, Reason is the correct explanation.'),
      option('b', 'Both are true but Reason is NOT the correct explanation.'),
      option('c', 'Assertion is true but Reason is false.'),
      option('d', 'Assertion is false but Reason is true.'),
      answer('(A) Both Assertion and Reason are true and Reason is the correct explanation.'),
      blank(),

      qPara(8, 'Assertion-Reason Question:'),
      assertion('A cook wondering why dal spilled out of a cooker is following a scientific method.'),
      reason('Science only happens in a laboratory.'),
      option('a', 'Both Assertion and Reason are true, Reason is the correct explanation.'),
      option('b', 'Both are true but Reason is NOT the correct explanation.'),
      option('c', 'Assertion is true but Reason is false.'),
      option('d', 'Assertion is false but Reason is true.'),
      answer('(C) Assertion is true but Reason is false.'),
      blank(),

      // ── Section E: Short Answer ─────────────────────────────────────────────
      new Paragraph({ text: 'E. SHORT ANSWER QUESTIONS (SA)', heading: HeadingLevel.HEADING_2 }),

      qPara(9, 'Give two examples of everyday professionals who apply scientific thinking.'),
      answer('Cook, bicycle repair person, electrician.'),
      blank(),

      qPara(10, 'What should you do if you cannot find an answer to a scientific question by yourself?'),
      answer('Ask your friends or teachers for help — scientists often work together to find answers.'),
      blank(),

      // ── Section F: Long Answer ──────────────────────────────────────────────
      new Paragraph({ text: 'F. LONG ANSWER QUESTION (LA)', heading: HeadingLevel.HEADING_2 }),

      qPara(11, 'Using the example of a bicycle repair person or an electrician, describe the steps they take that mirror the scientific method.'),
      answerLabel(),
      answerLine('They observe the problem (flat tyre or broken light).'),
      answerLine('They guess a possible cause (leak location or broken switch).'),
      answerLine('They test their idea (water bowl or checking the bulb).'),
      answerLine('They analyze whether their solution worked.'),
      blank(),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/sample_worksheet.docx', buffer);
  console.log('✅ sample_worksheet.docx created');
});
