const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, ExternalHyperlink } = require("docx");

const NAME_FONT = "Georgia";
const BODY_FONT = "Arial";

function hr() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "333333" } },
    spacing: { after: 120 },
  });
}

function sectionHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 22, font: BODY_FONT, color: "1A1A1A" })],
    spacing: { before: 140, after: 35 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" } },
  });
}

function jobHeader(left, right) {
  return new Paragraph({
    tabStops: [{ type: "right", position: 10540 }],
    children: [
      new TextRun({ text: left, bold: true, size: 21, font: BODY_FONT }),
      new TextRun({ text: "\t" + right, italics: true, size: 19, font: BODY_FONT }),
    ],
    spacing: { after: 20 },
  });
}

function projectHeader(left, url) {
  return new Paragraph({
    tabStops: [{ type: "right", position: 10540 }],
    children: [
      new TextRun({ text: left, bold: true, size: 21, font: BODY_FONT }),
      new TextRun({ text: "\t" }),
      new ExternalHyperlink({
        link: url,
        children: [new TextRun({ text: "View", size: 19, font: BODY_FONT, color: "000000" })],
      }),
    ],
    spacing: { after: 10 },
  });
}

function subLine(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 19, font: BODY_FONT })],
    spacing: { after: 25 },
  });
}

function bullet(text, boldParts) {
  // boldParts: array of substrings to bold within text (simple sequential split)
  const runs = [];
  let remaining = text;
  boldParts.forEach((b) => {
    const idx = remaining.indexOf(b);
    if (idx === -1) return;
    if (idx > 0) runs.push(new TextRun({ text: remaining.slice(0, idx), size: 19, font: BODY_FONT }));
    runs.push(new TextRun({ text: b, bold: true, size: 19, font: BODY_FONT }));
    remaining = remaining.slice(idx + b.length);
  });
  runs.push(new TextRun({ text: remaining, size: 19, font: BODY_FONT }));
  return new Paragraph({
    children: runs,
    bullet: { level: 0 },
    spacing: { after: 30 },
  });
}

function skillRow(label, value) {
  return new Paragraph({
    tabStops: [{ type: "left", position: 2900 }],
    indent: { left: 2900, hanging: 2900 },
    children: [
      new TextRun({ text: label, bold: true, size: 19, font: BODY_FONT }),
      new TextRun({ text: "\t" + value, size: 19, font: BODY_FONT }),
    ],
    spacing: { after: 40 },
  });
}

const doc = new Document({
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 }, // US Letter
          margin: { top: 260, bottom: 260, left: 850, right: 850 },
        },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "MAHATHI SABBANI", bold: true, size: 34, font: NAME_FONT })],
          spacing: { after: 60 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "Hyderabad, Telangana, India  |  +91 6303894983  |  mahasabbani16@gmail.com",
              size: 19,
              font: BODY_FONT,
            }),
          ],
          spacing: { after: 20 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [
            new ExternalHyperlink({
              link: "https://linkedin.com/in/mahathi-sabbani-932982317",
              children: [new TextRun({ text: "LinkedIn", size: 19, font: BODY_FONT, color: "000000" })],
            }),
            new TextRun({ text: "  \u25C7  ", size: 19, font: BODY_FONT, color: "333333" }),
            new ExternalHyperlink({
              link: "https://github.com/Mahathi1601",
              children: [new TextRun({ text: "GitHub", size: 19, font: BODY_FONT, color: "000000" })],
            }),
            new TextRun({ text: "  \u25C7  ", size: 19, font: BODY_FONT, color: "333333" }),
            new ExternalHyperlink({
              link: "https://leetcode.com/u/Mahathi_Sabbani/",
              children: [new TextRun({ text: "LeetCode", size: 19, font: BODY_FONT, color: "000000" })],
            }),
            new TextRun({ text: "  \u25C7  ", size: 19, font: BODY_FONT, color: "333333" }),
            new ExternalHyperlink({
              link: "https://hackerrank.com/profile/mahasabbani16",
              children: [new TextRun({ text: "HackerRank", size: 19, font: BODY_FONT, color: "000000" })],
            }),
            new TextRun({ text: "  \u25C7  ", size: 19, font: BODY_FONT, color: "333333" }),
            new ExternalHyperlink({
              link: "https://codechef.com/users/mahathi_16/",
              children: [new TextRun({ text: "CodeChef", size: 19, font: BODY_FONT, color: "000000" })],
            }),
          ],
        }),

        // EDUCATION
        sectionHeading("EDUCATION"),
        jobHeader("Gokaraju Rangaraju Institute of Engineering and Technology", "Hyderabad, India"),
        subLine("B.Tech in Computer Science Engineering (AI & ML) | Currently Studying — Expected Jul 2028"),
        jobHeader("Narayana Junior College", "Hyderabad, India"),
        subLine("Intermediate (MPC) | Percentage: 98% — Mar 2024"),
        jobHeader("Saint Claire High School", "Ramagundam, India"),
        subLine("CGPA: 9.7 — May 2022"),

        // EXPERIENCE
        sectionHeading("EXPERIENCE"),
        jobHeader("Amazon — Software Development Engineer Intern", "Jun 2026 – Aug 2026"),
        subLine("Central Shopping Experience Team | AWS Bedrock, fal.ai, Remotion, Rekognition | Bengaluru, India"),
        bullet(
          "Built an AI-powered video generation pipeline that converts an Amazon ASIN into a rendered product highlight video using AWS Bedrock (Claude), fal.ai, and Remotion.",
          ["AI-powered video generation pipeline", "AWS Bedrock (Claude), fal.ai, and Remotion"]
        ),
        bullet(
          "Automated storyboard creation and customer review curation with AWS Rekognition-based face filtering, improving relevance and reliability of generated videos at scale.",
          ["Automated storyboard creation and customer review curation", "AWS Rekognition"]
        ),
        jobHeader("Google for Developers — Android Developer Virtual Intern", "Jul 2025 – Sep 2025"),
        subLine("Supported by India Edu Program | Virtual"),
        bullet(
          "Completed a 10-week virtual internship in Android development, covering core frameworks and modern development tools.",
          ["10-week virtual internship"]
        ),
        bullet(
          "Gained hands-on experience building and debugging Android applications using industry-standard frameworks and tooling.",
          []
        ),

        // PROJECTS
        sectionHeading("PROJECTS"),
        projectHeader("Portfolio Website | HTML, CSS, JavaScript", "https://mahathi1601.github.io/Mahathi-personal-portfolio/"),
        bullet(
          "Designed and developed a personal portfolio website using HTML, CSS and JavaScript. Showcases projects, skills and resume in a structured and visually appealing way.",
          ["personal portfolio website"]
        ),
        projectHeader(
          "Artify: E-commerce Platform | MERN Stack, MongoDB Atlas, Render",
          "https://artify-2iq3.onrender.com/enter.html"
        ),
        bullet(
          "Developed and deployed a fully functional art store using the MERN stack with MongoDB Atlas and Render, enabling users to browse artworks, manage cart, and place orders with secure authentication.",
          ["fully functional art store", "browse artworks, manage cart, and place orders with secure authentication"]
        ),
        projectHeader(
          "Remote Interview Practice Portal | Node.js, MongoDB, Gemini AI",
          "https://remote-interview-practice-portal.vercel.app/"
        ),
        bullet(
          "Built a full-stack interview portal with OTP authentication, Gemini AI-powered interview evaluation, video mock recordings, and real-time analytics.",
          ["full-stack interview portal", "Gemini AI-powered interview evaluation"]
        ),
        projectHeader(
          "TerraGuard NER: Landslide Intelligence Platform | React, FastAPI, Gemini AI",
          "https://terra-guard-ner.vercel.app/"
        ),
        bullet(
          "Built an enterprise-grade landslide disaster intelligence and relief-routing platform for North-East India using React 18, Tailwind CSS, and FastAPI, integrating Google Gemini AI, OpenWeatherMap, and OpenRouteService for real-time hazard advisories and risk-aware convoy routing.",
          ["landslide disaster intelligence and relief-routing platform", "real-time hazard advisories and risk-aware convoy routing"]
        ),

        // SKILLS
        sectionHeading("SKILLS"),
        skillRow("Programming Languages:", "C, Java, JavaScript, Python"),
        skillRow("Web Technologies:", "HTML, CSS, React.js, Node.js, MERN Stack, Tailwind CSS, FastAPI"),
        skillRow("Databases:", "MongoDB, MongoDB Atlas"),
        skillRow("Cloud & AI Tools:", "AWS Bedrock, AWS Rekognition, fal.ai, Remotion, Gemini AI, OpenWeatherMap, OpenRouteService"),
        skillRow("Platforms:", "Android Development, Render, Vercel"),
        skillRow("Tools:", "Microsoft Excel, Git, GitHub"),

        // CERTIFICATIONS
        sectionHeading("CERTIFICATIONS"),
        bullet("Certification of Completion for Amazon Future Engineer Bootcamp on Java & DSA.", ["Amazon Future Engineer Bootcamp"]),
        bullet("Certification for the Completion of Linux Training (IIT Bombay).", ["Linux Training (IIT Bombay)"]),
        bullet("Certification for the Completion of Ruby Training (IIT Bombay).", ["Ruby Training (IIT Bombay)"]),
        bullet(
          'Certification of participation for National Level Online Workshop on "Application Development with AI & Essential Skills".',
          ["National Level Online Workshop"]
        ),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  require("fs").writeFileSync("Mahathi_Sabbani_Resume_ATS.docx", buffer);
  console.log("done");
});
