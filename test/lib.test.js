const { colLetter, buildMonthList, sumIfsFormula } = require("../src/lib.js");

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
