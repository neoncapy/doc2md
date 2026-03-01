# Contributing to Claude Code Orchestration Protocol

Thank you for your interest in improving this protocol. Contributions of all kinds are welcome: bug reports, feature suggestions, documentation improvements, and protocol refinements.

## Reporting Issues

Open a [GitHub issue](https://github.com/neoncapy/claude-code-orchestration-protocol/issues) with:

- A clear description of the problem
- The section of the protocol affected (e.g., QC loop, anti-polling, distillation)
- Your Claude Code version and model
- Steps to reproduce, if applicable

## Suggesting Improvements

The best protocol improvements come from real-world usage. If you have found a pattern that works well (or one that fails consistently), open an issue or pull request with:

- What you tried
- What happened
- What you expected
- Your proposed change

## Pull Request Process

1. Fork the repository
2. Create a feature branch (`git checkout -b improvement/description`)
3. Make your changes
4. Ensure all files use consistent formatting (Markdown, semantic XML tags)
5. Submit a pull request with a clear description

### Style Guidelines

- Protocol sections use semantic XML tags (`<section_name>`, not Markdown headings)
- Reference files use Markdown headings for readability
- Keep rules actionable and specific - avoid vague guidance
- Include WHY explanations for non-obvious rules
- Use full absolute paths in examples (not relative)
- Test your changes with actual Claude Code sessions when possible

### What Makes a Good Contribution

- Fixes for rules that cause agent failures in practice
- New anti-patterns discovered through real usage
- Improvements to the QC loop or distillation process
- Better examples in reference files
- Documentation clarifications

## Code of Conduct

This project follows the [Contributor Covenant](CODE_OF_CONDUCT.md). Please read it before participating.

## Questions?

Open a discussion issue. There are no bad questions when it comes to agent orchestration.
