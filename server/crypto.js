import crypto from "crypto";

const ENCRYPTION_ALGORITHM = "aes-256-gcm";
const TOKEN_KEY_LENGTH_BYTES = 32;

function tokenEncryptionKey() {
  const key = process.env.TOKEN_ENCRYPTION_KEY;

  if (!key) {
    throw new Error("TOKEN_ENCRYPTION_KEY is required for secure token storage.");
  }

  if (/^[a-f0-9]{64}$/i.test(key)) {
    return Buffer.from(key, "hex");
  }

  return crypto.createHash("sha256").update(key).digest().subarray(0, TOKEN_KEY_LENGTH_BYTES);
}

export function encryptSecret(value) {
  if (!value) return "";

  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv(ENCRYPTION_ALGORITHM, tokenEncryptionKey(), iv);
  const encrypted = Buffer.concat([cipher.update(String(value), "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();

  return [
    "v1",
    iv.toString("base64url"),
    tag.toString("base64url"),
    encrypted.toString("base64url"),
  ].join(".");
}

export function decryptSecret(value) {
  if (!value) return "";

  const [version, ivValue, tagValue, encryptedValue] = String(value).split(".");

  if (version !== "v1" || !ivValue || !tagValue || !encryptedValue) {
    throw new Error("Stored token format is invalid.");
  }

  const decipher = crypto.createDecipheriv(
    ENCRYPTION_ALGORITHM,
    tokenEncryptionKey(),
    Buffer.from(ivValue, "base64url")
  );
  decipher.setAuthTag(Buffer.from(tagValue, "base64url"));

  return Buffer.concat([
    decipher.update(Buffer.from(encryptedValue, "base64url")),
    decipher.final(),
  ]).toString("utf8");
}
