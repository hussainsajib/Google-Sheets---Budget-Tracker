const { colLetter, buildMonthList, sumIfsFormula, applyAutoCatRules, parseCsv, autoDetectMapping, normalizeDateString } = require("../src/lib.js");

describe("colLetter", () => {
  test("single-letter columns", () => {
    expect(colLetter(1)).toBe("A");
    expect(colLetter(2)).toBe("B");
    expect(colLetter(26)).toBe("Z");
  });
  test("double-letter columns", () => {
    expect(colLetter(27)).toBe("AA");
    expect(colLetter(28)).toBe("AB");
    expect(colLetter(52)).toBe("AZ");
    expect(colLetter(53)).toBe("BA");
  });
});

describe("buildMonthList", () => {
  test("returns the requested number of months", () => {
    expect(buildMonthList(2026, 1, 12)).toHaveLength(12);
    expect(buildMonthList(2026, 5, 3)).toHaveLength(3);
  });

  test("labels and month numbers are correct mid-year", () => {
    const m = buildMonthList(2026, 5, 3);
    expect(m[0]).toEqual({ year: 2026, month: 5, label: "May-26" });
    expect(m[1]).toEqual({ year: 2026, month: 6, label: "Jun-26" });
    expect(m[2]).toEqual({ year: 2026, month: 7, label: "Jul-26" });
  });

  test("rolls over into the next year", () => {
    const m = buildMonthList(2026, 11, 4); // Nov, Dec, Jan, Feb
    expect(m.map(x => x.label)).toEqual(["Nov-26", "Dec-26", "Jan-27", "Feb-27"]);
    expect(m[2]).toEqual({ year: 2027, month: 1, label: "Jan-27" });
  });

  test("handles a 12-month period starting in December", () => {
    const m = buildMonthList(2026, 12, 12);
    expect(m[0].label).toBe("Dec-26");
    expect(m[11]).toEqual({ year: 2027, month: 11, label: "Nov-27" });
  });
});

describe("applyAutoCatRules", () => {
  const rules = [
    { keyword: "uber",       category: "Transport"  },
    { keyword: "whole foods",category: "Groceries"  },
    { keyword: "netflix",    category: "Streaming"  },
  ];

  test("returns null when rules array is empty", () => {
    expect(applyAutoCatRules("Uber trip", [])).toBeNull();
  });

  test("returns null when rules is undefined/null", () => {
    expect(applyAutoCatRules("Uber trip", null)).toBeNull();
    expect(applyAutoCatRules("Uber trip", undefined)).toBeNull();
  });

  test("returns null when no rule matches", () => {
    expect(applyAutoCatRules("random merchant XYZ", rules)).toBeNull();
  });

  test("matches case-insensitively", () => {
    expect(applyAutoCatRules("UBER TRIP", rules)).toBe("Transport");
    expect(applyAutoCatRules("Uber Trip", rules)).toBe("Transport");
    expect(applyAutoCatRules("uber eats", rules)).toBe("Transport");
  });

  test("matches substring within a longer description", () => {
    expect(applyAutoCatRules("Payment to Uber Technologies Inc", rules)).toBe("Transport");
  });

  test("matches multi-word keyword", () => {
    expect(applyAutoCatRules("Whole Foods Market #321", rules)).toBe("Groceries");
  });

  test("returns the first matching rule when multiple rules could match", () => {
    const overlapping = [
      { keyword: "amazon",       category: "Shopping"  },
      { keyword: "amazon prime", category: "Streaming" },
    ];
    expect(applyAutoCatRules("Amazon Prime Video charge", overlapping)).toBe("Shopping");
  });

  test("skips rules with empty keywords", () => {
    const withEmpty = [
      { keyword: "",        category: "Shopping"  },
      { keyword: "netflix", category: "Streaming" },
    ];
    expect(applyAutoCatRules("netflix subscription", withEmpty)).toBe("Streaming");
  });

  test("skips rules with missing category", () => {
    const noCat = [
      { keyword: "uber", category: "" },
      { keyword: "uber", category: "Transport" },
    ];
    // first rule matches keyword but has no category — second rule wins
    expect(applyAutoCatRules("uber trip", noCat)).toBe("Transport");
  });

  test("empty description matches nothing", () => {
    expect(applyAutoCatRules("", rules)).toBeNull();
  });
});

describe("parseCsv", () => {
  test("parses a simple CSV with header and data rows", () => {
    const rows = parseCsv("Date,Description,Amount\n2026-01-01,Coffee,4.50\n2026-01-02,Bus,2.00");
    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual(["Date", "Description", "Amount"]);
    expect(rows[1]).toEqual(["2026-01-01", "Coffee", "4.50"]);
  });

  test("handles quoted fields containing commas", () => {
    const rows = parseCsv('Date,Description,Amount\n2026-01-01,"Smith, John",12.00');
    expect(rows[1][1]).toBe("Smith, John");
  });

  test("handles escaped double quotes inside quoted fields", () => {
    const rows = parseCsv('A,B\n1,"say ""hello"""');
    expect(rows[1][1]).toBe('say "hello"');
  });

  test("handles CRLF line endings", () => {
    const rows = parseCsv("A,B\r\n1,2\r\n3,4");
    expect(rows).toHaveLength(3);
    expect(rows[1]).toEqual(["1", "2"]);
  });

  test("strips UTF-8 BOM from start of input", () => {
    const rows = parseCsv("﻿Date,Amount\n2026-01-01,10");
    expect(rows[0][0]).toBe("Date");
  });

  test("drops empty trailing rows", () => {
    const rows = parseCsv("A,B\n1,2\n\n");
    expect(rows).toHaveLength(2);
  });

  test("returns empty array for blank input", () => {
    expect(parseCsv("")).toHaveLength(0);
    expect(parseCsv("   \n\n")).toHaveLength(0);
  });
});

describe("autoDetectMapping", () => {
  test("detects common bank CSV headers", () => {
    const m = autoDetectMapping(["Date", "Description", "Amount", "Account"]);
    expect(m.date).toBe(0);
    expect(m.description).toBe(1);
    expect(m.amount).toBe(2);
    expect(m.account).toBe(3);
  });

  test("returns -1 for unrecognised fields", () => {
    const m = autoDetectMapping(["Foo", "Bar"]);
    expect(m.date).toBe(-1);
    expect(m.amount).toBe(-1);
  });

  test("matches case-insensitively", () => {
    const m = autoDetectMapping(["POSTED DATE", "MEMO", "DEBIT"]);
    expect(m.date).toBe(0);
    expect(m.description).toBe(1);
    expect(m.amount).toBe(2);
  });

  test("maps 'Payee' to description", () => {
    const m = autoDetectMapping(["Trans Date", "Payee", "Amount"]);
    expect(m.description).toBe(1);
  });

  test("maps 'Category' header", () => {
    const m = autoDetectMapping(["Date", "Merchant", "Category", "Amount"]);
    expect(m.category).toBe(2);
  });

  test("first match wins; later columns are not double-mapped", () => {
    const m = autoDetectMapping(["Date", "Posted Date", "Amount"]);
    expect(m.date).toBe(0);
  });
});

describe("normalizeDateString", () => {
  test("passes through ISO yyyy-MM-dd unchanged", () => {
    expect(normalizeDateString("2026-06-10", "YMD")).toBe("2026-06-10");
  });

  test("parses yyyy/MM/dd", () => {
    expect(normalizeDateString("2026/06/10", "YMD")).toBe("2026-06-10");
  });

  test("parses yyyyMMdd compact", () => {
    expect(normalizeDateString("20260610", "MDY")).toBe("2026-06-10");
  });

  test("parses MM/DD/YYYY (MDY format)", () => {
    expect(normalizeDateString("06/10/2026", "MDY")).toBe("2026-06-10");
  });

  test("parses DD/MM/YYYY (DMY format)", () => {
    expect(normalizeDateString("10/06/2026", "DMY")).toBe("2026-06-10");
  });

  test("parses 2-digit year MM/DD/YY", () => {
    expect(normalizeDateString("06/10/26", "MDY")).toBe("2026-06-10");
  });

  test("parses d-Mon-yyyy", () => {
    expect(normalizeDateString("10-Jun-2026", "MDY")).toBe("2026-06-10");
  });

  test("parses Mon d, yyyy", () => {
    expect(normalizeDateString("Jun 10, 2026", "MDY")).toBe("2026-06-10");
  });

  test("parses d Mon yyyy (no comma)", () => {
    expect(normalizeDateString("10 Jun 2026", "MDY")).toBe("2026-06-10");
  });

  test("returns empty string for empty input", () => {
    expect(normalizeDateString("", "MDY")).toBe("");
  });

  test("returns original string for unrecognised format", () => {
    expect(normalizeDateString("not-a-date", "MDY")).toBe("not-a-date");
  });
});

describe("sumIfsFormula", () => {
  test("builds a correct month-bounded SUMIFS", () => {
    const f = sumIfsFormula("$A12", { year: 2026, month: 7 });
    expect(f).toContain("SUMIFS(Transactions!$D:$D");
    expect(f).toContain("Transactions!$C:$C,$A12");
    expect(f).toContain('">="&DATE(2026,7,1)');
    expect(f).toContain('"<="&EOMONTH(DATE(2026,7,1),0)');
    expect(f.startsWith("=IFERROR(")).toBe(true);
  });
});
