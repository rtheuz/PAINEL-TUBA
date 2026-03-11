You are a **senior software engineer specialized in Google Apps Script architecture and performance optimization**.

Your task is to **analyze and refactor a large Google Apps Script codebase (~7000 lines)** while preserving **100% of its current behavior and functionality**.

IMPORTANT CONSTRAINTS:

1. **Do NOT remove any existing functionality.**
2. **Do NOT change external behavior or outputs.**
3. The optimized code must remain fully compatible with:

   * Google Apps Script runtime
   * Google Sheets API via SpreadsheetApp
   * Google Drive API via DriveApp
4. Maintain all current data structures and business logic.

OPTIMIZATION GOALS:

1. **Performance optimization**

   * Reduce calls to SpreadsheetApp.getRange()
   * Minimize calls to getDataRange()
   * Batch read/write operations
   * Replace appendRow() with setValues() where appropriate
   * Avoid unnecessary SpreadsheetApp.flush()

2. **Code architecture**

   * Modularize code into logical sections:

     * data access
     * business logic
     * drive operations
     * utilities
   * Extract reusable helpers
   * Reduce duplication

3. **Memory and runtime efficiency**

   * Cache spreadsheet data when possible
   * Avoid repeated sheet lookups
   * Use lazy initialization patterns

4. **Maintainability**

   * Improve naming clarity
   * Reduce function complexity
   * Add concise comments explaining critical logic
   * Keep the code readable and structured

5. **Safety**

   * Preserve all sheet column mappings
   * Preserve all folder structures and naming conventions
   * Preserve all calculations and pricing logic

OUTPUT FORMAT:

1. First provide a **technical analysis**:

   * performance bottlenecks
   * architectural issues
   * potential improvements

2. Then provide the **optimized version of the code**.

3. Clearly mark:

   * modified sections
   * new helper functions
   * performance improvements

4. If the file is too large, refactor it into multiple logical modules but ensure **the final system still works exactly the same**.

CONTEXT ABOUT THE SYSTEM:

The script is a production system that:

* manages product catalog
* calculates metal fabrication quotes
* generates project folders in Google Drive
* manages clients and suppliers
* generates proposals and production orders
* interacts heavily with Google Sheets formulas

Do not simplify the system.
Only **optimize and restructure it while preserving all existing behavior**.

I will now provide the Google Apps Script code for analysis and optimization.
