const { showPrompt, getTournamentNameParsed } = require("../utils");

test("showPrompt should return user response", () => {
  // Mock the SpreadsheetApp and Logger
  global.SpreadsheetApp = {
    getUi: jest.fn().mockReturnValue({
      prompt: jest.fn().mockReturnValue({
        getSelectedButton: jest.fn().mockReturnValue("OK"),
        getResponseText: jest.fn().mockReturnValue("Test Response"),
      }),
      ButtonSet: {
        OK_CANCEL: "OK_CANCEL",
      },
      Button: {
        OK: "OK",
      },
    }),
  };
  global.Logger = {
    log: jest.fn(),
  };

  const response = showPrompt("Test Prompt");
  expect(response).toBe("Test Response");
});

test("getTournamentNameParsed should return tournament name", () => {
  // Mock the SpreadsheetApp and Utilities
  global.SpreadsheetApp = {
    getActiveSpreadsheet: jest.fn().mockReturnValue({
      getRangeByName: jest.fn().mockReturnValue({
        getValue: jest.fn().mockReturnValue("Test Tournament"),
      }),
    }),
  };
  global.Utilities = {
    formatDate: jest.fn().mockReturnValue("1-January-2023"),
  };

  const tournamentName = getTournamentNameParsed();
  expect(tournamentName).toBe(
    "1-January-2023 Test Tournament Division-Test Tournament @ Test Tournament",
  );
});
