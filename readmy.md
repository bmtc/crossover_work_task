# Vendor Spend Strategy Assessment

## Overview
This project analyzes vendor spend data for an acquired company to identify major cost-saving opportunities. The analysis targets a $1B+ business and requires detailed vendor categorization, strategic recommendations, and executive-level summary.

## Project Structure
```
.
├── input/
│   └── input_file.xlsx          # Raw vendor spend data (~400 vendors, 12-month history)
├── output/
│   └── output_file.xlsx         # Completed analysis with all vendor classifications
├── scripts/
│   └── [Python scripts for analysis automation]
├── readmy.md                    # This file
├── prompts.txt                  # All prompts used during analysis
└── Task description.docx        # Original assessment requirements
```

## Task Requirements

### Part 1: Analyze Vendor Data
For each vendor in the spreadsheet, complete:
- **Department**: Classify vendor (e.g., Engineering, G&A, Finance, Support, etc.)
- **Vendor Description**: One-line concise description of vendor purpose
- **Strategic Recommendation**: Choose one action
  - **Terminate**: Vendor no longer needed
  - **Consolidate**: Multiple vendors serve same function; streamline
  - **Optimize**: Useful vendor with cost/usage reduction opportunity

### Part 2: Identify Strategic Opportunities
Top 3 Opportunities tab:
- Three highest-impact recommendations
- Each includes: title, explanation, estimated annual savings (USD)

### Part 3: Summarize Methodology
Methodology tab:
- Approach to the task
- Tools used
- Prompts created
- Validation & quality-check methodology with evidence

### Part 4: Executive Memo
Recommendations tab - 1-page executive memo:
- Audience: CEO and CFO
- Summary of findings and savings opportunities
- Realistic timeline
- Implementation process
- Risk assessment
- Clear, actionable, aligned with C-level expectations

## Key Constraints
⚠️ **MUST use Claude Code CLI** - Not claude.ai or other AI tools
- All work must be automated/generated using Claude Code
- Vendor categorizations must be accurate and justified
- Descriptions must be concise and specific
- Recommendations must be realistic and strategic
- Savings must be significant enough to "move the needle" for a $1B+ business

## Quality Checklist
- [ ] All ~400 vendors accurately categorized into departments
- [ ] Vendor descriptions are concise, specific, and accurate
- [ ] Recommendations are realistic, strategic, and justified
- [ ] Risk factors identified
- [ ] Top 3 opportunities are specific, plausible, and financially justified
- [ ] Methodology clearly explained with tool usage documented
- [ ] Quality checks documented (AI or manual) with evidence
- [ ] Executive memo is well-formatted, error-free, actionable
- [ ] Savings are significant and realistic
- [ ] Project folder is well-organized with clear documentation

## Next Steps
1. ✅ Initialize GitHub repository
2. Generate automated analysis scripts
3. Categorize all vendors using Claude Code
4. Identify top opportunities and cost-saving potential
5. Quality check all outputs
6. Prepare executive memo
7. Final submission and validation
