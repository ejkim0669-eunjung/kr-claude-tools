---
name: find-skills
description: Use this skill whenever the user asks what skills are available, wants to know which skill to use for a task, asks "what can you do?", types /find-skills or /skills, asks to list or search skills, says they are not sure which tool or skill fits their need, or asks about Claude Code capabilities. Also trigger when the user mentions a task type (charting, PDF, code review, scheduling, etc.) and you want to suggest the right skill before diving in.
---

# Find Skills

Help the user discover installed skills and match them to tasks.

## When invoked

The user wants to know what skills exist, or you spotted a task where a skill would help.

## Steps

1. **Get the roster** - load `ListSkills` via ToolSearch and call it. This returns every installed skill with its name and description.

2. **Decide the display mode**

   - **Browse mode** (user asked "what skills are available?" or similar): show ALL skills as a compact table - name, one-line description.
   - **Match mode** (user described a task): identify the 1-3 best-fit skills and explain briefly why each fits. Skip ones that clearly do not apply.

3. **Format rules**
   - Use a Markdown table or tight bullet list, not paragraphs.
   - Skill name as a slash command: `/dataviz`, `/code-review`, etc.
   - One line per skill; two lines max for the best match in match mode.
   - End with: "Type `/skill-name` to activate, or ask me to use it."

4. **If ListSkills returns nothing** - say: "No skills are installed yet. Skills give Claude specialized knowledge for specific tasks. Browse available skills in Claude Code settings -> Skills."

## Example output (browse mode)

| Skill | What it does |
|---|---|
| `/dataviz` | Charts, graphs, color palettes, dashboards |
| `/code-review` | Finds correctness bugs and suggests fixes |
| `/artifact-design` | Design guidance before writing any artifact |
| `/deep-research` | Multi-step web research with sourced findings |

## Example output (match mode)

User: "I want to make a sales chart from my CSV"

> For charting, use **`/dataviz`** - it covers chart type selection, color palettes, and layout for any visualization. Just say `/dataviz` and describe what you want to show.

## Notes

- Skills load instructions into context, so they add tokens upfront but reduce back-and-forth. Worth it for recurring, specialized tasks.
- You do not need to load a skill to describe it - the roster description is enough.
- Never invent skill names. Only list what `ListSkills` actually returns.
