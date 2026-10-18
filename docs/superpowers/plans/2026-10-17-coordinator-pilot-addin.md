# Coordinator pilot add-in plan

> Execute inline using superpowers:executing-plans. User approved the proposed XLAM/Russian Ribbon/installation stage and delegates implementation details.

**Goal:** locally reviewable coordinator pilot ZIP containing AccessToolsv2.xlam, Russian Ribbon, matching package metadata, artificial example and installation instructions.

**Architecture:** build tools import only src into an owned Excel instance, compile, save XLAM and inject an Open XML Ribbon. Existing commands/checks stay intact except a small package-consistency guard. No startup code, source indexing or automatic installation. Package guard compares embedded version/release/profile/commit with installed metadata before commands. SHA256 manifest is independently verifiable by build/installation tooling, not a claimed cryptographic runtime guard.

**Scope:** coordinator-only pilot, not official production release. Full synchronized FIRST_LINE/COORDINATOR release builder and published Admin configuration remain outside this pilot. No user security settings changed and no automatic installation/merge/GitHub release.

## Tasks

- [x] RED/GREEN VBA package metadata checks; guard both Russian commands; add two Ribbon callbacks and static XML.
- [x] RED/GREEN Ribbon ZIP injection tool, preserving all existing entries and rejecting duplicate parts/relations. Tests use synthetic ZIP.
- [x] Build supervisor runs existing static/full Excel gate, then bounded owned worker imports src only (strict ANSI), embeds pilot constants, compiles and saves XLAM. Package gets metadata/config/instructions/artificial inputs and manifest; ZIP created only after validation. Existing output path rejected.
- [x] Verify packaged file format/add-in state, no test modules, callbacks, clean load/no input books at startup, metadata mismatch rejection, and actual run+save on artificial input from a clean install copy. UI click/dialog acceptance explicitly separate if unavailable.
- [x] Full regression, tooling, independent review, clean copy, draft PR against PR26. Deliver local ZIP/file with clear pilot limitations.

## Risks/acceptance

Never weaken macro security or install globally. Only owned Excel processes closed; supervisor kills only PID/start-time matching process on timeout. No user sources in package or repo. New release directory required, no overwrites. Excel reload after Ribbon injection necessary. Initial Ribbon is COORDINATOR-only; no misleading FIRST_LINE artifact. Constant DEV bypass exists only in source test host and must be replaced/verified by builder. Metadata mismatch blocks UI commands; no claim of resistance to malicious tampering.

Evidence: full/root+cleanbuild11308/0, packaged-copy11/0, extracted-ZIP11/0, all file SHA unchanged, static7/0 gate9/0 Ribbon2/0, fresh review clean. Source08422b5. UI registration/buttons/dialogs remain manual acceptance, not claimed complete.

18.10 user override: v2 filename AccessToolsv2.xlam for parallel loading with first-version AccessTools.xlam; new pilot.2 build. Existing legacy references in smoke are deliberate synthetic coexistence fixture, not runtime v2 paths.
