# QA Test Suite - Academic Email Assistant

## Overview

This directory contains the complete test suite for the Academic Email Assistant Outlook add-in.

## Test Structure

```
tests/
├── README.md              # This file
├── setup.js               # Jest setup with global mocks
├── fixtures/
│   └── sampleEmails.json  # 10 synthetic academic emails for testing
├── unit/
│   ├── utilityFunctions.test.js   # Tests for hashString, extractMessages, extractText
│   ├── themeDetection.test.js     # Tests for applyTheme() logic
│   ├── tokenStorage.test.js       # Tests for localStorage token persistence
│   └── rpcLayer.test.js           # Tests for RPC call/response handling
└── manual/
    └── TEST_PROCEDURES.md # Step-by-step manual test procedures
```

## Running Tests

### Unit Tests
```bash
npm test              # Run all tests
npm run test:unit     # Run unit tests only
npm run test:coverage # Run with coverage report
```

### Manual Test Execution
All manual tests are documented in `manual/TEST_PROCEDURES.md`. These require:
- Microsoft Outlook Desktop with sideloaded Academic Email Assistant
- OpenClaw Gateway running on localhost:18789
- Active LLM connection (Ollama)

## Test Coverage

**60 unit tests** covering:
- Utility functions (hashString, extractMessages, extractText, isRawToolCall)
- Theme detection logic (light/dark mode)
- Token storage and persistence
- RPC layer (callRpc, handleIncoming)
- Event handling

## Test Data

10 synthetic academic emails in `fixtures/sampleEmails.json` covering:
- Email-001: Assignment deadline inquiry (medium urgency, inquiry)
- Email-002: Exam marks access issue (high urgency, complaint)
- Email-003: Workshop attendance (low urgency, administrative)
- Email-004: Extension request (medium urgency, request)
- Email-005: Lecture feedback (low urgency, feedback)
- Email-006: Exam format question (medium urgency, inquiry)
- Email-007: Course materials follow-up (low urgency, inquiry)
- Email-008: Meeting request (medium urgency, request)
- Email-009: Library resources announcement (low urgency, administrative)
- Email-010: Group project complaint (high urgency, complaint)

## Manual Test Cases

The manual test procedures are documented in `manual/TEST_PROCEDURES.md` and cover both response-based testing and system/integration testing.

### System, UI, and Integration Tests
- TC-01: Outlook add-in loading
- TC-02: Email context extraction
- TC-03: Gateway token connection
- TC-04: Invalid or missing token handling
- TC-05: End-to-end workflow connection validation
- TC-09: Empty/unreadable email body
- TC-11: Very long email body
- TC-12: Gateway/model unavailable
- TC-14: Privacy and data handling

### Response-Based and Extended Manual Tests
- TC-06: Basic email question
- TC-07: Draft reply generation
- TC-08: Use Draft in Outlook
- TC-10: Ambiguous email content
- TC-13: Per-email chat history
- TC-15: Performance and code coverage (manual timing required)
- TC-16: Pinned sidebar / email switching
- TC-17: Token persistence
- TC-18: Theme compatibility

## Notes

The unit tests test **logic in isolation** by defining local copies of functions to test. This approach was chosen because the actual source code (taskpane.js) depends heavily on Office.js browser APIs that require complex mocking. The tests validate the correctness of the algorithmic logic.

For full integration testing with Outlook, run the manual test procedures documented in `manual/TEST_PROCEDURES.md`.
