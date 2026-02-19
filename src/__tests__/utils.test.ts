import { describe, it, expect } from "vitest";
import { ok, err } from "../utils.js";

describe("ok", () => {
  it("wraps an object as pretty-printed JSON text content", () => {
    const result = ok({ id: "abc", title: "Task" });
    expect(result.content).toHaveLength(1);
    expect(result.content[0].type).toBe("text");
    expect(JSON.parse(result.content[0].text)).toEqual({ id: "abc", title: "Task" });
  });

  it("wraps an array", () => {
    const result = ok([1, 2, 3]);
    expect(JSON.parse(result.content[0].text)).toEqual([1, 2, 3]);
  });

  it("wraps null", () => {
    const result = ok(null);
    expect(result.content[0].text).toBe("null");
  });

  it("wraps a primitive string", () => {
    const result = ok("hello");
    expect(JSON.parse(result.content[0].text)).toBe("hello");
  });

  it("wraps a boolean", () => {
    const result = ok(true);
    expect(JSON.parse(result.content[0].text)).toBe(true);
  });
});

describe("err", () => {
  it("extracts the message from an Error instance", () => {
    const result = err(new Error("something went wrong"));
    expect(result.content[0].text).toBe("Error: something went wrong");
    expect(result.isError).toBe(true);
  });

  it("converts a plain string to an error message", () => {
    const result = err("oops");
    expect(result.content[0].text).toBe("Error: oops");
    expect(result.isError).toBe(true);
  });

  it("converts an unknown object via String()", () => {
    const result = err({ code: 404 });
    expect(result.content[0].text).toBe("Error: [object Object]");
    expect(result.isError).toBe(true);
  });

  it("converts null", () => {
    const result = err(null);
    expect(result.content[0].text).toBe("Error: null");
    expect(result.isError).toBe(true);
  });

  it("converts undefined", () => {
    const result = err(undefined);
    expect(result.content[0].text).toBe("Error: undefined");
    expect(result.isError).toBe(true);
  });
});
