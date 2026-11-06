type IconName =
  | "location"
  | "phone"
  | "mail"
  | "linkedin"
  | "print"
  | "word"

function Icon({ name }: { name: IconName }) {
  const paths: Record<IconName, React.ReactNode> = {
    location: (
      <>
        <path d="M20 10c0 5-8 12-8 12S4 15 4 10a8 8 0 1 1 16 0Z" />
        <circle cx="12" cy="10" r="2.5" />
      </>
    ),
    phone: (
      <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.8 19.8 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.12 4.18 2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.36 1.9.69 2.8a2 2 0 0 1-.45 2.11L8.08 9.9a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.9.33 1.84.56 2.8.69A2 2 0 0 1 22 16.92Z" />
    ),
    mail: (
      <>
        <rect width="20" height="16" x="2" y="4" rx="2" />
        <path d="m22 7-10 6L2 7" />
      </>
    ),
    linkedin: (
      <>
        <path d="M16 8a6 6 0 0 1 6 6v7h-4v-7a2 2 0 0 0-4 0v7h-4v-7a6 6 0 0 1 6-6Z" />
        <path d="M2 9h4v12H2z" />
        <circle cx="4" cy="4" r="2" />
      </>
    ),
    print: (
      <>
        <path d="M6 9V2h12v7" />
        <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" />
        <rect width="12" height="8" x="6" y="14" />
      </>
    ),
    word: (
      <>
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z" />
        <path d="M14 2v6h6" />
        <path d="m8 13 1.3 5 1.7-4 1.7 4 1.3-5" />
      </>
    ),
  }

  return (
    <svg
      aria-hidden="true"
      className="icon"
      fill="none"
      viewBox="0 0 24 24"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      {paths[name]}
    </svg>
  )
}

const competencies = [
  "Financial Analysis",
  "Excel",
  "Market Research",
  "Data Analytics",
  "Budgeting",
  "Client Presentations",
  "Program Operations",
  "Team Leadership",
  "Stakeholder Communication",
  "Content Strategy",
  "Competitive Analysis",
  "SEO",
]

const experience = [
  {
    dates: "2026 — Present",
    company: "Ross School of Business Advising Office · Ann Arbor, MI",
    role: "Ross Peer Mentor",
    highlights: [
      "Selected and trained to mentor five first-year Ross students on academics, campus adjustment, and extracurricular opportunities.",
      "Guided a five-student consulting case for Underground Printing, assigning roles and developing SEO recommendations.",
      "Connected mentees with academic resources and student organizations aligned with their interests.",
    ],
  },
  {
    dates: "2025 — Present",
    company: "Michigan Innovative Marketing · Ann Arbor, MI",
    role: "Marketing Consultant",
    highlights: [
      "Led analytics for a five-person team during a six-week engagement with an Ann Arbor fashion retailer.",
      "Surveyed 50 students and benchmarked five competitors to identify campus awareness gaps and content opportunities.",
      "Built a content calendar with employee ownership and six performance metrics to standardize execution and measurement.",
    ],
  },
  {
    dates: "2025 — Present",
    company: "Phi Gamma Nu Professional Business Fraternity · Ann Arbor, MI",
    role: "Member & Wellness Chair",
    highlights: [
      "Presented a short thesis on Klarna Group using delinquency trends and repayment-loss exposure to support the recommendation.",
      "Completed 80 coffee chats, Excel analysis, and industry presentations through the chapter’s professional-development curriculum.",
      "Led two fundraising initiatives, engaging 20+ participants and raising $500+ for a local homeless shelter.",
    ],
  },
  {
    dates: "2025 — Present",
    company: "Rochester Avon Recreation Authority · Rochester, MI",
    role: "Owner, RARA Golf",
    highlights: [
      "Founded a youth golf instruction program serving 60+ students and generating $13K in revenue across 17 programs.",
      "Recruited and trained four instructors while coordinating staffing, class expectations, pricing, compensation, equipment, and expenses.",
      "Maintained consistent profitability while building students’ skills, confidence, and long-term enjoyment of the game.",
    ],
  },
  {
    dates: "2021 — Present",
    company: "Twin Lakes Golf & Swim Club · Oakland Township, MI",
    role: "Assistant Golf Instructor",
    highlights: [
      "Led 18 youth camps, including three all-girl sessions, using skill-based drills and competitions for groups of eight.",
      "Helped 7+ golfers ages 6–13 reduce average nine-hole scores by three strokes through swing and course-management coaching.",
      "Coached a women’s league through tournament preparation with personalized instruction in fundamentals, strategy, and etiquette.",
    ],
  },
]

async function exportToWord() {
  const {
    AlignmentType,
    BorderStyle,
    Document,
    Packer,
    Paragraph,
    TextRun,
  } = await import("docx")

  const sectionHeading = (text: string) =>
    new Paragraph({
      spacing: { before: 240, after: 120 },
      border: {
        bottom: {
          color: "FFCB05",
          style: BorderStyle.SINGLE,
          size: 12,
          space: 5,
        },
      },
      children: [
        new TextRun({
          text: text.toUpperCase(),
          bold: true,
          color: "00274C",
          size: 22,
        }),
      ],
    })

  const experienceParagraphs = experience.flatMap((item) => [
    new Paragraph({
      spacing: { before: 180, after: 20 },
      children: [
        new TextRun({
          text: item.role,
          bold: true,
          color: "00274C",
          size: 23,
        }),
        new TextRun({
          text: `  |  ${item.dates}`,
          bold: true,
          color: "69798A",
          size: 18,
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 70 },
      children: [
        new TextRun({
          text: item.company,
          bold: true,
          color: "315D7C",
          size: 18,
        }),
      ],
    }),
    ...item.highlights.map(
      (highlight) =>
        new Paragraph({
          text: highlight,
          bullet: { level: 0 },
          spacing: { after: 45, line: 260 },
        }),
    ),
  ])

  const wordDocument = new Document({
    creator: "Madison Yang",
    title: "Madison Yang Resume",
    description:
      "Michigan Ross BBA sophomore resume for internship opportunities.",
    styles: {
      default: {
        document: {
          run: { font: "Aptos", size: 20, color: "394B5F" },
          paragraph: { spacing: { line: 276 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 540, right: 650, bottom: 540, left: 650 },
          },
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 70 },
            children: [
              new TextRun({
                text: "MADISON YANG",
                bold: true,
                color: "00274C",
                size: 38,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 55 },
            children: [
              new TextRun({
                text: "MICHIGAN ROSS BBA SOPHOMORE  |  FINANCE EMPHASIS",
                bold: true,
                color: "315D7C",
                size: 19,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 160 },
            children: [
              new TextRun(
                "Ann Arbor, MI  •  248.525.5837  •  madiyang@umich.edu  •  linkedin.com/in/yang-madison",
              ),
            ],
          }),
          sectionHeading("Profile"),
          new Paragraph({
            spacing: { after: 80, line: 280 },
            children: [
              new TextRun(
                "Michigan Ross sophomore combining entrepreneurial ownership, financial discipline, and analytical problem-solving across youth programming, marketing consulting, and student leadership. ",
              ),
              new TextRun({
                text: "Brings a record of turning insights into action—from building a profitable golf business to delivering data-backed client recommendations.",
                bold: true,
              }),
            ],
          }),
          sectionHeading("Education"),
          new Paragraph({
            spacing: { after: 35 },
            children: [
              new TextRun({
                text: "University of Michigan — Stephen M. Ross School of Business",
                bold: true,
                color: "00274C",
                size: 22,
              }),
            ],
          }),
          new Paragraph({
            text: "Bachelor of Business Administration, Finance Emphasis  |  Expected May 2029",
            spacing: { after: 35 },
          }),
          new Paragraph({
            text: "GPA: 3.5 / 4.0  |  University of Michigan Regents Scholarship",
          }),
          sectionHeading("Experience & Leadership"),
          ...experienceParagraphs,
          sectionHeading("Core Competencies"),
          new Paragraph({
            text: competencies.join("  •  "),
            spacing: { after: 80, line: 280 },
          }),
          sectionHeading("Additional"),
          new Paragraph({
            text: "Michigan Club Golf Team — 0.0 GHIN; 4th at Michigan State Finals",
            bullet: { level: 0 },
            spacing: { after: 45 },
          }),
          new Paragraph({
            text: "Interests: global travel, creative arts, weightlifting, reformer Pilates, and hot yoga",
            bullet: { level: 0 },
          }),
        ],
      },
    ],
  })

  const blob = await Packer.toBlob(wordDocument)
  const url = URL.createObjectURL(blob)
  const link = document.createElement("a")
  link.href = url
  link.download = "Madison-Yang-Resume.docx"
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
}

function SectionTitle({ children }: { children: React.ReactNode }) {
  return (
    <div className="section-heading">
      <span>{children}</span>
      <span className="section-rule" />
    </div>
  )
}

export default function App() {
  return (
    <main className="page-shell">
      <div className="action-buttons" aria-label="Resume export options">
        <button
          className="word-button"
          type="button"
          onClick={() => void exportToWord()}
          aria-label="Export resume to Microsoft Word"
        >
          <Icon name="word" />
          <span>Export Word</span>
        </button>
        <button
          className="print-button"
          type="button"
          onClick={() => window.print()}
          aria-label="Print or save resume as PDF"
        >
          <Icon name="print" />
          <span>Print / Save PDF</span>
        </button>
      </div>

      <article className="cv">
        <header className="hero">
          <div className="hero-portrait">
            <img
              src={`${import.meta.env.BASE_URL}madison-yang-headshot.jpg`}
              alt="Madison Yang"
              width="821"
              height="900"
            />
          </div>
          <div className="hero-copy">
            <p className="eyebrow">Michigan Ross</p>
            <h1>Madison Yang</h1>
            <p className="title">BBA Sophomore · Finance Emphasis</p>
            <p className="specialization">
              Finance <span>|</span> Strategy <span>|</span> Analytics
            </p>
          </div>
          <div className="contact-list" aria-label="Contact details">
            <div>
              <Icon name="location" />
              <span>Ann Arbor, Michigan</span>
            </div>
            <a href="tel:+12485255837">
              <Icon name="phone" />
              <span>248.525.5837</span>
            </a>
            <a href="mailto:madiyang@umich.edu">
              <Icon name="mail" />
              <span>madiyang@umich.edu</span>
            </a>
            <a
              href="https://linkedin.com/in/yang-madison"
              target="_blank"
              rel="noreferrer"
            >
              <Icon name="linkedin" />
              <span>linkedin.com/in/yang-madison</span>
            </a>
          </div>
        </header>

        <div className="cv-grid">
          <aside className="sidebar">
            <section>
              <SectionTitle>At a glance</SectionTitle>
              <div className="stats">
                <div className="stat">
                  <strong>
                    60<span>+</span>
                  </strong>
                  <p>Students served</p>
                </div>
                <div className="stat">
                  <strong>
                    $13<span>K</span>
                  </strong>
                  <p>Program revenue</p>
                </div>
              </div>
              <ul className="check-list">
                <li>Entrepreneur with hands-on P&amp;L ownership</li>
                <li>Consulting, analytics, and client-facing experience</li>
                <li>Team leader across business, mentoring, and athletics</li>
              </ul>
            </section>

            <section>
              <SectionTitle>Core competencies</SectionTitle>
              <div className="skill-list">
                {competencies.map((competency) => (
                  <span key={competency}>{competency}</span>
                ))}
              </div>
            </section>

            <section>
              <SectionTitle>Education</SectionTitle>
              <div className="education-item">
                <img
                  className="education-logo"
                  src={`${import.meta.env.BASE_URL}university-of-michigan.svg`}
                  alt=""
                  aria-hidden="true"
                />
                <div className="education-copy">
                  <p className="degree">BBA</p>
                  <p>
                    Finance Emphasis
                    <br />
                    Expected May 2029
                  </p>
                  <strong>Michigan Ross</strong>
                </div>
              </div>
              <div className="education-details">
                <p>
                  <span>GPA</span>
                  <strong>3.5 / 4.0</strong>
                </p>
                <p>
                  <span>Honor</span>
                  <strong>Regents Scholarship</strong>
                </p>
              </div>
            </section>

            <section>
              <SectionTitle>Leadership &amp; interests</SectionTitle>
              <ul className="detail-list">
                <li>
                  <strong>Michigan Club Golf</strong>
                  <span>0.0 GHIN · 4th at Michigan State Finals</span>
                </li>
                <li>
                  <strong>Global Travel</strong>
                  <span>Food and culture across Serbia and Japan</span>
                </li>
                <li>
                  <strong>Creative Arts</strong>
                  <span>Nail design, eyelashes, and hairstyling</span>
                </li>
                <li>
                  <strong>Fitness</strong>
                  <span>Weightlifting, reformer Pilates, and hot yoga</span>
                </li>
              </ul>
            </section>
          </aside>

          <div className="main-content">
            <section className="summary">
              <SectionTitle>Profile</SectionTitle>
              <p>
                Michigan Ross sophomore combining entrepreneurial ownership,
                financial discipline, and analytical problem-solving across
                youth programming, marketing consulting, and student leadership.
              </p>
              <p>
                Brings a record of turning insights into action—from building a
                profitable golf business to delivering data-backed client
                recommendations—and is eager to contribute that drive in a
                high-impact internship.
              </p>
            </section>

            <section>
              <SectionTitle>Experience &amp; leadership</SectionTitle>
              <div className="timeline">
                {experience.map((item) => (
                  <article
                    className="role"
                    key={`${item.company}-${item.role}`}
                  >
                    <div className="timeline-marker" />
                    <p className="dates">{item.dates}</p>
                    <div className="role-copy">
                      <p className="company">{item.company}</p>
                      <h3>{item.role}</h3>
                      <ul className="role-highlights">
                        {item.highlights.map((highlight) => (
                          <li key={highlight}>{highlight}</li>
                        ))}
                      </ul>
                    </div>
                  </article>
                ))}
              </div>
            </section>
          </div>
        </div>

        <footer>
          <span>Madison Yang</span>
          <span>Michigan Ross · BBA 2029</span>
        </footer>
      </article>
    </main>
  )
}
