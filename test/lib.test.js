const { colLetter, buildMonthList, sumIfsFormula, applyAutoCatRules } = require("../src/lib.js");

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
