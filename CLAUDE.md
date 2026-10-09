# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Current state

This repository is a fresh scaffold. As of the initial commit it contains only a
one-line `README.md` with the project name and no source code, build tooling,
tests, or CI configuration. The name suggests a tool for exporting Highcharts
chart data into Excel workbooks, but nothing has been implemented yet.

Because there is no code, there are no build, lint, or test commands to
document. Do not assume a language, package manager, or framework; check the
repository root for a manifest (`package.json`, `pyproject.toml`, etc.) before
running anything.

## Keeping this file useful

Once the project takes shape, update this file with:

- The commands a contributor actually runs: install, build, lint, test, and
  how to run a single test.
- The high-level architecture that is not obvious from the file tree, such as
  how Highcharts series/option data is read, how it is mapped to worksheet
  cells, and where Excel output is produced.
- Any repository-specific conventions that differ from the defaults of the
  chosen toolchain.

Avoid listing every file or restating generic best practices.
