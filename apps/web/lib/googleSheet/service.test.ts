import { afterEach, beforeEach, describe, expect, test, vi } from "vitest";
import { AuthenticationError, OperationNotAllowedError, UnknownError } from "@formbricks/types/errors";
import { TIntegrationGoogleSheets } from "@formbricks/types/integration/google-sheet";
import {
  GOOGLE_SHEET_INTEGRATION_INSUFFICIENT_PERMISSION,
  GOOGLE_SHEET_INTEGRATION_INSUFFICIENT_SCOPES,
  GOOGLE_SHEET_INTEGRATION_INVALID_GRANT,
} from "@/lib/googleSheet/constants";

vi.mock("@/lib/integration/service", () => ({ createOrUpdateIntegration: vi.fn() }));

vi.mock("@/lib/constants", () => ({
  GOOGLE_SHEETS_CLIENT_ID: "client-id",
  GOOGLE_SHEETS_CLIENT_SECRET: "client-secret",
  GOOGLE_SHEETS_REDIRECT_URL: "https://example.com/callback",
  GOOGLE_SHEET_MESSAGE_LIMIT: 50000,
}));

const sheetsMock = vi.hoisted(() => ({
  update: vi.fn(),
  append: vi.fn(),
}));

const spreadsheetsGetMock = vi.hoisted(() => vi.fn());

vi.mock("googleapis", () => ({
  google: {
    auth: {
      OAuth2: class {
        setCredentials() {}
      },
    },
    sheets: () => ({ spreadsheets: { values: sheetsMock, get: spreadsheetsGetMock } }),
  },
}));

const { getSpreadsheetNameById, writeData } = await import("@/lib/googleSheet/service");

/**
 * The shape the integration has once ENG-2078's redaction has blanked `config.key`: schema-valid,
 * because `ZGoogleCredential` types both tokens as `z.string()`, but unusable for auth.
 */
const redactedIntegration = {
  id: "integration1",
  type: "googleSheets",
  workspaceId: "ws1",
  config: {
    email: "owner@example.com",
    data: [],
    key: {
      scope: "https://www.googleapis.com/auth/spreadsheets",
      token_type: "Bearer",
      expiry_date: 0,
      access_token: "",
      refresh_token: "",
    },
  },
} as TIntegrationGoogleSheets;

/** An integration whose stored access token is still valid, so `authorize` never has to refresh. */
const authorizedIntegration = {
  id: "integration1",
  type: "googleSheets",
  workspaceId: "ws1",
  config: {
    email: "owner@example.com",
    data: [],
    key: {
      scope: "https://www.googleapis.com/auth/spreadsheets",
      token_type: "Bearer",
      expiry_date: Date.now() + 60 * 60 * 1000,
      access_token: "access-token",
      refresh_token: "refresh-token",
    },
  },
} as TIntegrationGoogleSheets;

// ENG-2303: a blank refresh token reached googleapis, which surfaced it as a bare "No refresh token is
// set." straight into a user-facing toast. It has to fail as the reconnect case instead, which the modal
// maps to a real message. Nothing here touches the network: the guard short-circuits before any Google
// call, which is also what keeps this test hermetic.
describe("getSpreadsheetNameById", () => {
  test("throws an invalid_grant AuthenticationError when the stored refresh token is blank", async () => {
    await expect(getSpreadsheetNameById(redactedIntegration, "sheet1")).rejects.toThrow(
      new AuthenticationError(GOOGLE_SHEET_INTEGRATION_INVALID_GRANT)
    );
  });

  // ENG-2807: Google's scope rejection never says "permission", so it slipped past the permission guard
  // and reached the user as a raw English Google string. It needs its own code, because the fix is to
  // reconnect and grant the Sheets scope, not to share the spreadsheet.
  describe("when Google rejects the spreadsheet fetch", () => {
    const rejectWith = (message: string) => {
      spreadsheetsGetMock.mockImplementation((_params, callback: (err: Error | null) => void) => {
        setTimeout(() => callback(new Error(message)), 0);
      });
    };

    beforeEach(() => {
      spreadsheetsGetMock.mockReset();
      vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: true }));
    });

    afterEach(() => {
      vi.unstubAllGlobals();
    });

    test("maps a missing OAuth scope to the insufficient_scopes code", async () => {
      rejectWith("Request had insufficient authentication scopes.");

      await expect(getSpreadsheetNameById(authorizedIntegration, "sheet1")).rejects.toThrow(
        new OperationNotAllowedError(GOOGLE_SHEET_INTEGRATION_INSUFFICIENT_SCOPES)
      );
    });

    test("keeps mapping a sharing problem to the insufficient_permission code", async () => {
      rejectWith("The caller does not have permission");

      await expect(getSpreadsheetNameById(authorizedIntegration, "sheet1")).rejects.toThrow(
        new OperationNotAllowedError(GOOGLE_SHEET_INTEGRATION_INSUFFICIENT_PERMISSION)
      );
    });

    test("wraps any other Google error in an UnknownError", async () => {
      rejectWith("Requested entity was not found.");

      await expect(getSpreadsheetNameById(authorizedIntegration, "sheet1")).rejects.toThrow(UnknownError);
    });
  });
});

// ENG-2250: both writes used to `throw` from inside the node-style googleapis callback, so the error
// escaped `writeData`'s own `try/catch` and surfaced as an unhandled rejection while `writeData` itself
// resolved — the pipeline recorded a success and the customer's sheet quietly stopped filling up.
describe("writeData", () => {
  /**
   * googleapis never calls back on the same tick, and that is the whole bug: a synchronous callback
   * would carry the old `throw` back into `writeData`'s `try/catch` and hide it. Every mock here defers
   * like the real client does.
   */
  const respondAsync =
    (err: Error | null, onCall?: () => void) => (_params: unknown, callback: (err: Error | null) => void) => {
      setTimeout(() => {
        onCall?.();
        callback(err);
      }, 0);
    };

  beforeEach(() => {
    sheetsMock.update.mockReset();
    sheetsMock.append.mockReset();
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: true }));
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  test("rejects with an UnknownError when the append callback reports an error", async () => {
    sheetsMock.update.mockImplementation(respondAsync(null));
    sheetsMock.append.mockImplementation(respondAsync(new Error("The caller does not have permission")));

    await expect(writeData(authorizedIntegration, "sheet1", ["answer"], ["question"])).rejects.toThrow(
      new UnknownError("Error while appending data: The caller does not have permission")
    );
  });

  test("rejects with an UnknownError when the header update callback reports an error", async () => {
    sheetsMock.update.mockImplementation(respondAsync(new Error("Requested entity was not found")));

    await expect(writeData(authorizedIntegration, "sheet1", ["answer"], ["question"])).rejects.toThrow(
      new UnknownError("Error while appending data: Requested entity was not found")
    );
    expect(sheetsMock.append).not.toHaveBeenCalled();
  });

  test("resolves only after both writes have completed", async () => {
    const completed: string[] = [];
    sheetsMock.update.mockImplementation(respondAsync(null, () => completed.push("update")));
    sheetsMock.append.mockImplementation(respondAsync(null, () => completed.push("append")));

    await writeData(authorizedIntegration, "sheet1", ["answer"], ["question"]);

    expect(completed).toEqual(["update", "append"]);
  });

  test("appends distinct rows when the first response cell is empty and preserves existing rows", async () => {
    const sheet: string[][] = [
      ["Statement", "Answer"],
      ["", "earlier response"],
    ];

    // Test assumption, not a verified implementation of Google's table detection: a search
    // limited to column A misses this table when A2 is blank and writes at A2 again. A range
    // spanning response columns finds the existing row by its populated later cells.
    const simulateAppendPlacement = (target: string[][], range: string, row: string[]) => {
      const searchOnlyFirstColumn = range === "A2";
      const lastMatchingRow = target.reduce((last, existingRow, index) => {
        if (index === 0) return last; // The append search starts below the header row.
        const hasDataInSearchRange = searchOnlyFirstColumn
          ? existingRow[0] !== ""
          : existingRow.some((cell) => cell !== "");
        return hasDataInSearchRange ? index : last;
      }, 0);
      const destinationRow = searchOnlyFirstColumn && lastMatchingRow === 0 ? 1 : lastMatchingRow + 1;
      target[destinationRow] = row;
    };

    // Explicitly show the regression represented by the assumption: the second old-range
    // write targets A2 again and replaces the first response row.
    const oldRangeSheet = structuredClone(sheet);
    simulateAppendPlacement(oldRangeSheet, "A2", ["", "first new response"]);
    simulateAppendPlacement(oldRangeSheet, "A2", ["", "second new response"]);
    expect(oldRangeSheet).toEqual([
      ["Statement", "Answer"],
      ["", "second new response"],
    ]);

    sheetsMock.update.mockImplementation(
      (
        params: { range: string; requestBody: { values: string[][] } },
        callback: (err: Error | null) => void
      ) => {
        setTimeout(() => {
          if (params.range === "A1") sheet[0] = params.requestBody.values[0];
          callback(null);
        }, 0);
      }
    );
    sheetsMock.append.mockImplementation(
      (
        params: { range: string; requestBody: { values: string[][] } },
        callback: (err: Error | null) => void
      ) => {
        setTimeout(() => {
          simulateAppendPlacement(sheet, params.range, params.requestBody.values[0]);
          callback(null);
        }, 0);
      }
    );

    await writeData(authorizedIntegration, "sheet1", ["", "first new response"], ["Statement", "Answer"]);
    await writeData(authorizedIntegration, "sheet1", ["", "second new response"], ["Statement", "Answer"]);

    expect(sheet).toEqual([
      ["Statement", "Answer"],
      ["", "earlier response"],
      ["", "first new response"],
      ["", "second new response"],
    ]);
    expect(sheetsMock.append.mock.calls.map(([params]) => params.range)).toEqual(["A2:ZZ", "A2:ZZ"]);
    expect(sheetsMock.append).toHaveBeenCalledTimes(2);
    expect(sheetsMock.append.mock.calls.map(([params]) => params.requestBody.values[0])).toEqual([
      ["", "first new response"],
      ["", "second new response"],
    ]);
  });

  test("continues appending responses whose first cell is non-empty", async () => {
    const sheet: string[][] = [["Statement", "Answer"]];
    sheetsMock.update.mockImplementation(
      (
        _params: { range: string; requestBody: { values: string[][] } },
        callback: (err: Error | null) => void
      ) => {
        setTimeout(() => callback(null), 0);
      }
    );
    sheetsMock.append.mockImplementation(
      (
        params: { range: string; requestBody: { values: string[][] } },
        callback: (err: Error | null) => void
      ) => {
        setTimeout(() => {
          sheet.push(params.requestBody.values[0]);
          callback(null);
        }, 0);
      }
    );

    await writeData(authorizedIntegration, "sheet1", ["first", "answer 1"], ["Statement", "Answer"]);
    await writeData(authorizedIntegration, "sheet1", ["second", "answer 2"], ["Statement", "Answer"]);

    expect(sheet).toEqual([
      ["Statement", "Answer"],
      ["first", "answer 1"],
      ["second", "answer 2"],
    ]);
  });
});
