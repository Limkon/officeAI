import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import selfsigned from 'selfsigned';

export interface CertPair {
  key: string;
  cert: string;
  keyPath: string;
  certPath: string;
}

function isCertValid(certPem: string): boolean {
  try {
    const x509 = new crypto.X509Certificate(certPem);
    const expiresAt = Date.parse(x509.validTo);
    if (Number.isNaN(expiresAt)) return false;
    // Renew if expiring within 30 days
    const thirtyDaysMs = 30 * 24 * 60 * 60 * 1000;
    return expiresAt - Date.now() > thirtyDaysMs;
  } catch {
    return false;
  }
}

export function ensureLocalhostCert(dataDir: string): CertPair {
  const certDir = path.join(dataDir, 'certs');
  fs.mkdirSync(certDir, { recursive: true });
  const keyPath = path.join(certDir, 'localhost.key');
  const certPath = path.join(certDir, 'localhost.crt');

  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    const key = fs.readFileSync(keyPath, 'utf8');
    const cert = fs.readFileSync(certPath, 'utf8');
    if (key.trim() && isCertValid(cert)) {
      return { key, cert, keyPath, certPath };
    }
  }

  const attrs = [{ name: 'commonName', value: 'OfficeAI Local Gateway' }];
  const pems = selfsigned.generate(attrs, {
    keySize: 2048,
    days: 820,
    algorithm: 'sha256',
    extensions: [
      {
        name: 'basicConstraints',
        cA: true
      },
      {
        name: 'keyUsage',
        keyCertSign: true,
        digitalSignature: true,
        nonRepudiation: true,
        keyEncipherment: true,
        dataEncipherment: true
      },
      {
        name: 'extKeyUsage',
        serverAuth: true
      },
      {
        name: 'subjectAltName',
        altNames: [
          { type: 2, value: 'localhost' },
          { type: 7, ip: '127.0.0.1' }
        ]
      }
    ]
  });

  fs.writeFileSync(keyPath, pems.private, 'utf8');
  fs.writeFileSync(certPath, pems.cert, 'utf8');

  return {
    key: pems.private,
    cert: pems.cert,
    keyPath,
    certPath
  };
}
