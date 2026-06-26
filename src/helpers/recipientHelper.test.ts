import { isExternalEmail, hasExternalRecipients, getRecipientEmails } from "./recipientHelper";

describe("recipientHelper - isExternalEmail", () => {
  const companyDomain = "@aatb.com.br";

  test("should identify internal emails correctly", () => {
    expect(isExternalEmail("test@aatb.com.br", companyDomain)).toBe(false);
    expect(isExternalEmail("TEST@AATB.COM.BR", companyDomain)).toBe(false);
    expect(isExternalEmail("user.name@aatb.com.br", companyDomain)).toBe(false);
  });

  test("should identify external emails correctly", () => {
    expect(isExternalEmail("test@gmail.com", companyDomain)).toBe(true);
    expect(isExternalEmail("test@outlook.com", companyDomain)).toBe(true);
    expect(isExternalEmail("test@aatb.com.br.other.com", companyDomain)).toBe(true);
  });

  test("should handle domain without leading @ symbol", () => {
    const domainWithoutAt = "aatb.com.br";
    expect(isExternalEmail("test@aatb.com.br", domainWithoutAt)).toBe(false);
    expect(isExternalEmail("test@gmail.com", domainWithoutAt)).toBe(true);
  });

  test("should handle empty or null values gracefully", () => {
    expect(isExternalEmail("", companyDomain)).toBe(false);
    expect(isExternalEmail(null as any, companyDomain)).toBe(false);
  });
});

describe("recipientHelper - hasExternalRecipients", () => {
  const companyDomain = "@aatb.com.br";

  test("should return false if all emails are internal", () => {
    const emails = ["user1@aatb.com.br", "user2@aatb.com.br"];
    expect(hasExternalRecipients(emails, companyDomain)).toBe(false);
  });

  test("should return true if at least one email is external", () => {
    const emails = ["user1@aatb.com.br", "external@gmail.com", "user2@aatb.com.br"];
    expect(hasExternalRecipients(emails, companyDomain)).toBe(true);
  });

  test("should return false for empty lists", () => {
    expect(hasExternalRecipients([], companyDomain)).toBe(false);
  });
});

describe("recipientHelper - getRecipientEmails", () => {
  test("should fetch and extract email addresses from all fields successfully", async () => {
    const mockItem = {
      to: {
        getAsync: jest.fn((callback) => {
          callback({
            status: "succeeded",
            value: [
              { emailAddress: "to1@aatb.com.br" },
              { emailAddress: "to2@aatb.com.br" }
            ]
          });
        })
      },
      cc: {
        getAsync: jest.fn((callback) => {
          callback({
            status: "succeeded",
            value: [
              { emailAddress: "cc1@aatb.com.br" }
            ]
          });
        })
      },
      bcc: {
        getAsync: jest.fn((callback) => {
          callback({
            status: "succeeded",
            value: [
              { emailAddress: "bcc1@gmail.com" }
            ]
          });
        })
      }
    };

    const emails = await getRecipientEmails(mockItem);
    expect(emails).toEqual([
      "to1@aatb.com.br",
      "to2@aatb.com.br",
      "cc1@aatb.com.br",
      "bcc1@gmail.com"
    ]);

    expect(mockItem.to.getAsync).toHaveBeenCalledTimes(1);
    expect(mockItem.cc.getAsync).toHaveBeenCalledTimes(1);
    expect(mockItem.bcc.getAsync).toHaveBeenCalledTimes(1);
  });

  test("should ignore failed or empty recipient fields gracefully", async () => {
    const mockItem = {
      to: {
        getAsync: jest.fn((callback) => {
          callback({
            status: "succeeded",
            value: [{ emailAddress: "to1@aatb.com.br" }]
          });
        })
      },
      cc: {
        getAsync: jest.fn((callback) => {
          callback({
            status: "failed",
            error: { message: "Some error occurred" }
          });
        })
      }
      // bcc is undefined
    };

    const emails = await getRecipientEmails(mockItem);
    expect(emails).toEqual(["to1@aatb.com.br"]);
  });

  test("should return empty list if item is null or undefined", async () => {
    expect(await getRecipientEmails(null)).toEqual([]);
    expect(await getRecipientEmails(undefined)).toEqual([]);
  });
});
