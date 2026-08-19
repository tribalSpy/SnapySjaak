import crypto from "node:crypto";

// End-to-end sealed-box encryption for backup payloads: X25519 ECDH (per-file
// ephemeral key) -> HKDF-SHA256 -> AES-256-GCM. Render only ever holds the
// public key (BACKUP_PUBLIC_KEY) and can seal but never open; only the office
// PC holds BACKUP_PRIVATE_KEY, so a leak of Render's own env vars can never
// decrypt a backup, past or present.

const AES_ALGORITHM = "aes-256-gcm";
const HKDF_INFO = Buffer.from("snappysjaak-backup-v1");
const IV_LENGTH = 12;
const AUTH_TAG_LENGTH = 16;

export function generateBackupKeyPair() {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("x25519");
  return {
    publicKey: publicKey.export({ type: "spki", format: "der" }).toString("base64"),
    privateKey: privateKey.export({ type: "pkcs8", format: "der" }).toString("base64"),
  };
}

function loadPublicKey(base64Value) {
  return crypto.createPublicKey({ key: Buffer.from(String(base64Value || ""), "base64"), format: "der", type: "spki" });
}

function loadPrivateKey(base64Value) {
  return crypto.createPrivateKey({ key: Buffer.from(String(base64Value || ""), "base64"), format: "der", type: "pkcs8" });
}

function deriveAesKey(sharedSecret) {
  return Buffer.from(crypto.hkdfSync("sha256", sharedSecret, Buffer.alloc(0), HKDF_INFO, 32));
}

// Envelope layout: [2-byte BE ephemeral-pubkey length][ephemeral pubkey DER][12-byte IV][16-byte auth tag][ciphertext]
export function sealBuffer(plaintext, recipientPublicKeyBase64) {
  const recipientPublicKey = loadPublicKey(recipientPublicKeyBase64);
  const ephemeral = crypto.generateKeyPairSync("x25519");
  const sharedSecret = crypto.diffieHellman({ privateKey: ephemeral.privateKey, publicKey: recipientPublicKey });
  const aesKey = deriveAesKey(sharedSecret);
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(AES_ALGORITHM, aesKey, iv);
  const ciphertext = Buffer.concat([cipher.update(plaintext), cipher.final()]);
  const authTag = cipher.getAuthTag();
  const ephemeralPublicKeyDer = ephemeral.publicKey.export({ type: "spki", format: "der" });

  const header = Buffer.alloc(2);
  header.writeUInt16BE(ephemeralPublicKeyDer.length, 0);
  return Buffer.concat([header, ephemeralPublicKeyDer, iv, authTag, ciphertext]);
}

export function openSealed(envelope, recipientPrivateKeyBase64) {
  const recipientPrivateKey = loadPrivateKey(recipientPrivateKeyBase64);
  let offset = 0;
  const ephemeralPublicKeyLength = envelope.readUInt16BE(offset);
  offset += 2;
  const ephemeralPublicKeyDer = envelope.subarray(offset, offset + ephemeralPublicKeyLength);
  offset += ephemeralPublicKeyLength;
  const iv = envelope.subarray(offset, offset + IV_LENGTH);
  offset += IV_LENGTH;
  const authTag = envelope.subarray(offset, offset + AUTH_TAG_LENGTH);
  offset += AUTH_TAG_LENGTH;
  const ciphertext = envelope.subarray(offset);

  const ephemeralPublicKey = crypto.createPublicKey({ key: ephemeralPublicKeyDer, format: "der", type: "spki" });
  const sharedSecret = crypto.diffieHellman({ privateKey: recipientPrivateKey, publicKey: ephemeralPublicKey });
  const aesKey = deriveAesKey(sharedSecret);
  const decipher = crypto.createDecipheriv(AES_ALGORITHM, aesKey, iv);
  decipher.setAuthTag(authTag);
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]);
}
