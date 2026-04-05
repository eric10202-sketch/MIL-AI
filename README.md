# Carveout AI Toolkit

AI-assisted document generation for IT carve-out projects, powered by GitHub Copilot (GPT-5.4).

---

## Prerequisites

| Tool | Where to get it |
|------|----------------|
| [VS Code](https://code.visualstudio.com/) | code.visualstudio.com |
| GitHub Copilot extension | VS Code Extensions marketplace → search **GitHub Copilot** |
| Active Copilot licence | Confirm with your IT/licence admin |
| Git | [git-scm.com](https://git-scm.com/) |
| Bosch DevCloud account | Required to clone this repo |

---

## Setup (one-time, ~10 minutes)

**1. Clone the repository**

```
git clone https://github.boschdevcloud.com/kho1sgp/SUaaS
```

**2. Open the workspace in VS Code**

Open the `Carveout/` subfolder as your workspace:

- File → Open Folder → select `SUaaS/Carveout`

> Copilot will automatically load the workspace instructions from `.github/copilot-instructions.md`.

**3. Sign in to GitHub Copilot**

Click the Copilot icon in the VS Code status bar and sign in with your GitHub account.

**4. Enable Agent mode**

In the Copilot Chat panel, switch the mode dropdown to **Agent**.

---

## Starting a New Carve-out Project

Open Copilot Chat in Agent mode and provide the following details:

| Field | Example |
|-------|---------|
| Project name | Project Phoenix |
| Seller | Bosch |
| Buyer | Acme Corp |
| Business carved out | Power Tools Division |
| Carve-out model | Stand Alone / Integration / Combination |
| PMO lead | Your name |
| Worldwide sites | 12 |
| IT users | 3,500 |
| Project start date | 01 May 2026 |
| GoLive date | 01 Feb 2027 |
| Project completion date | 30 Apr 2027 |

Copilot will enforce the compliance gate, then generate all deliverables in the correct order:

1. Schedule (XLSX + MS Project XML)
2. Risk register (Excel using `BD_Risk-Register_template_en_V1.0_Dec2023.xlsx`)
3. Cost plan (XLSX)
4. Project charter (HTML)
5. Executive dashboard (HTML)
6. Management KPI dashboard (HTML)
7. Monthly status report (PDF)
8. Management stakeholder presentation (PPTX) — final executive brief
9. Multi-document reconciliation gate — final quality check before handoff

All newly created projects must use repository versioning:

1. Initial project creation = `v1.0`
2. Minor approved change = `v1.x`
3. Major approved change or rebaseline = `v2.0`, `v3.0`, and so on

For new work, store outputs in a versioned active-project folder such as `Project Phoenix v1.0`. Create a new versioned folder for later updates instead of overwriting the prior approved version.

---

## Saving Your Work

After Copilot generates files, commit and push them to keep the repo up to date:

```
git add .
git commit -m "Add [Project Name] initial deliverables"
git push origin main
```

---

## Folder Structure

```
Carveout/
├── .agents/skills/          # Active AI skill instructions (loaded on demand)
├── .github/                 # Copilot workspace instructions
├── AlphaX/                  # Example project — reference only
├── Falcon/                  # Example project — reference only
├── active-projects/         # Track active engagement folders here
├── archive/                 # Closed projects
├── references/              # Methodology reference materials
├── templates/               # Document templates
├── BD_Risk-Register_template_en_V1.0_Dec2023.xlsx
├── Bosch.png                # Logo used in generated documents
└── README.md                # This file
```

---

## Questions / Issues

Contact the toolkit owner or raise an issue in the [SUaaS repository](https://github.boschdevcloud.com/kho1sgp/SUaaS).
