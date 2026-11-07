/*****************************************************************
 * HUELLA (llaves de acceso / WebAuthn) para acciones de administrador
 * del tablero, como descartar fotos dudosas.
 *
 * La huella nunca llega aquí: el celular o el PC firma un reto con una
 * llave que solo se usa si la persona pone su dedo (o su PIN de Windows).
 * Registrar una huella nueva exige el CODIGO_ADMIN (Propiedades del script).
 * Viene de la prueba de huella que funcionó el 7 nov 2026.
 *****************************************************************/

const ORIGEN = 'https://geol21.github.io';   // dónde vive la PWA
const RPID = 'geol21.github.io';             // dominio al que queda atada la llave
const MINUTOS_RETO = 5;

/** Ejecútala desde el editor para comprobar la verificación de firmas (debe decir true · true). */
function probarHuella() {
  const r = verificarFirmaP256_(PRUEBA.x, PRUEBA.y, hex_(PRUEBA.hash), PRUEBA.r, PRUEBA.s);
  const mal = verificarFirmaP256_(PRUEBA.x, PRUEBA.y, hex_(PRUEBA.hash.replace(/^./, PRUEBA.hash[0] === 'a' ? 'b' : 'a')), PRUEBA.r, PRUEBA.s);
  Logger.log('Firma válida aceptada: ' + r + ' · Firma alterada rechazada: ' + !mal);
}

/** ¿Esta llave sigue registrada como administrador? */
function huellaRegistrada_(credId) {
  const g = PropertiesService.getScriptProperties().getProperty('HUELLA_' + String(credId || ''));
  return g ? JSON.parse(g).nombre : '';
}

/** Registrar huella de administrador: pide el CODIGO_ADMIN, con bloqueo tras varios intentos fallidos. */
function registrarHuellaAdmin_(d) {
  const props = PropertiesService.getScriptProperties();
  const admin = props.getProperty('CODIGO_ADMIN');
  if (!admin) throw new Error('Falta crear CODIGO_ADMIN en las Propiedades del script.');
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('ADMIN_FALLOS') || 0);
  if (fallos >= 10) throw new Error('Demasiados intentos. Espera 15 minutos.');
  if (String(d.codigoAdmin || '').trim() !== admin) {
    cache.put('ADMIN_FALLOS', String(fallos + 1), 900);
    throw new Error('Código de administrador incorrecto.');
  }
  cache.remove('ADMIN_FALLOS');
  return registrar_(d);
}

/***********************
 * RETO: un número al azar, válido 5 minutos y para un solo uso
 ***********************/

function reto_(d) {
  const bytes = [];
  for (let i = 0; i < 2; i++) Utilities.getUuid().replace(/-/g, '').match(/../g).forEach(function(h) { bytes.push(parseInt(h, 16)); });
  const reto = b64url_(bytes);
  const id = Utilities.getUuid();
  CacheService.getScriptCache().put('RETO_' + id, JSON.stringify({ reto: reto, para: String(d.para || '') }), MINUTOS_RETO * 60);
  return { ok: true, id: id, reto: reto };
}

function tomarReto_(id) {
  const c = CacheService.getScriptCache();
  const v = c.get('RETO_' + id);
  if (!v) throw new Error('El reto venció o ya se usó. Intenta de nuevo.');
  c.remove('RETO_' + id); // un solo uso
  return JSON.parse(v);
}

function revisarClientData_(b64, tipo, reto) {
  const cd = JSON.parse(texto_(desdeB64url_(b64)));
  if (cd.type !== tipo) throw new Error('Tipo de operación inesperado: ' + cd.type);
  if (cd.challenge !== reto.reto) throw new Error('El reto no coincide.');
  if (cd.origin !== ORIGEN) throw new Error('La firma viene de otra página: ' + cd.origin);
  return cd;
}


/***********************
 * REGISTRO: se guarda la llave pública (la privada nunca sale del celular)
 ***********************/

function registrar_(d) {
  const reto = tomarReto_(d.retoId);
  revisarClientData_(d.clientDataJSON, 'webauthn.create', reto);
  if (Number(d.alg) !== -7) throw new Error('Este celular usa un tipo de llave que la prueba no conoce (' + d.alg + ').');
  const spki = desdeB64url_(d.spki);
  const k = spki.slice(spki.length - 65);
  if (k[0] !== 4) throw new Error('Llave pública con formato inesperado.');
  const nombre = String(d.nombre || '').trim() || 'Sin nombre';
  const cred = { nombre: nombre, x: hexDe_(k.slice(1, 33)), y: hexDe_(k.slice(33, 65)), creada: new Date().toISOString() };
  if (!enCurva_(BigInt('0x' + cred.x), BigInt('0x' + cred.y))) throw new Error('La llave pública no es válida.');
  PropertiesService.getScriptProperties().setProperty('HUELLA_' + d.credId, JSON.stringify(cred));
  return { ok: true, nombre: nombre };
}

/***********************
 * VERIFICACIÓN: el celular firmó el reto con la huella
 ***********************/

function verificar_(d) {
  const reto = tomarReto_(d.retoId);
  const guardada = PropertiesService.getScriptProperties().getProperty('HUELLA_' + d.credId);
  if (!guardada) throw new Error('Esta huella no está registrada en el servidor.');
  const cred = JSON.parse(guardada);
  revisarClientData_(d.clientDataJSON, 'webauthn.get', reto);

  const auth = desdeB64url_(d.authenticatorData);
  const rpHash = sha256_(bytesDe_(RPID));
  for (let i = 0; i < 32; i++) if (auth[i] !== rpHash[i]) throw new Error('La llave es de otro sitio.');
  const flags = auth[32];
  if (!(flags & 0x01)) throw new Error('No se confirmó la presencia de la persona.');
  if (!(flags & 0x04)) throw new Error('El celular no verificó la huella o el bloqueo de pantalla.');

  // Lo firmado = authenticatorData + SHA-256(clientDataJSON)
  const firmado = auth.concat(sha256_(desdeB64url_(d.clientDataJSON)));
  const rs = derAFirma_(desdeB64url_(d.signature));
  const valido = verificarFirmaP256_(cred.x, cred.y, sha256_(firmado), rs.r, rs.s);
  if (!valido) throw new Error('La firma no es válida.');
  return { ok: true, nombre: cred.nombre, para: reto.para, hora: new Date().toISOString() };
}


/***********************
 * ECDSA P-256 (Apps Script no trae verificación de este tipo de firma)
 ***********************/

const P_ = BigInt('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff');
const A_ = P_ - BigInt(3);
const B_ = BigInt('0x5ac635d8aa3a93e7b3ebbd55769886bc651d06b0cc53b0f63bce3c3e27d2604b');
const N_ = BigInt('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551');
const G_ = [BigInt('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'),
            BigInt('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')];
const CERO_ = BigInt(0), UNO_ = BigInt(1), DOS_ = BigInt(2), TRES_ = BigInt(3);

function mod_(a, m) { const r = a % m; return r < CERO_ ? r + m : r; }

function inv_(a, m) {
  let t = CERO_, nt = UNO_, r = m, nr = mod_(a, m);
  while (nr !== CERO_) {
    const q = r / nr;
    let tmp = t - q * nt; t = nt; nt = tmp;
    tmp = r - q * nr; r = nr; nr = tmp;
  }
  if (r !== UNO_) throw new Error('Sin inverso.');
  return mod_(t, m);
}

function enCurva_(x, y) {
  if (x < CERO_ || x >= P_ || y < CERO_ || y >= P_) return false;
  return mod_(y * y - (x * x * x + A_ * x + B_), P_) === CERO_;
}

/** Suma de puntos en coordenadas afines; null = punto en el infinito. */
function sumar_(p, q) {
  if (!p) return q;
  if (!q) return p;
  if (p[0] === q[0]) {
    if (mod_(p[1] + q[1], P_) === CERO_) return null;
    const l = mod_((TRES_ * p[0] * p[0] + A_) * inv_(DOS_ * p[1], P_), P_);
    const x = mod_(l * l - DOS_ * p[0], P_);
    return [x, mod_(l * (p[0] - x) - p[1], P_)];
  }
  const l = mod_((q[1] - p[1]) * inv_(q[0] - p[0], P_), P_);
  const x = mod_(l * l - p[0] - q[0], P_);
  return [x, mod_(l * (p[0] - x) - p[1], P_)];
}

function multiplicar_(k, p) {
  let r = null, a = p;
  while (k > CERO_) {
    if (k & UNO_) r = sumar_(r, a);
    a = sumar_(a, a);
    k = k >> UNO_;
  }
  return r;
}

/** x, y en hexadecimal; hash como arreglo de bytes; r, s como BigInt o hex. */
function verificarFirmaP256_(xHex, yHex, hash, r, s) {
  r = typeof r === 'string' ? BigInt('0x' + r) : r;
  s = typeof s === 'string' ? BigInt('0x' + s) : s;
  if (r <= CERO_ || r >= N_ || s <= CERO_ || s >= N_) return false;
  const Q = [BigInt('0x' + xHex), BigInt('0x' + yHex)];
  if (!enCurva_(Q[0], Q[1])) return false;
  const e = BigInt('0x' + hexDe_(hash));
  const w = inv_(s, N_);
  const X = sumar_(multiplicar_(mod_(e * w, N_), G_), multiplicar_(mod_(r * w, N_), Q));
  return !!X && mod_(X[0], N_) === r;
}

/** Firma en formato DER (lo que entrega el celular) → r y s. */
function derAFirma_(b) {
  if (b[0] !== 0x30) throw new Error('Firma con formato inesperado.');
  let i = 2;
  if (b[i] !== 0x02) throw new Error('Firma con formato inesperado.');
  const lr = b[i + 1]; const r = b.slice(i + 2, i + 2 + lr); i = i + 2 + lr;
  if (b[i] !== 0x02) throw new Error('Firma con formato inesperado.');
  const ls = b[i + 1]; const s = b.slice(i + 2, i + 2 + ls);
  return { r: BigInt('0x' + hexDe_(r)), s: BigInt('0x' + hexDe_(s)) };
}


/***********************
 * BYTES (Apps Script usa bytes con signo: se pasan a 0–255)
 ***********************/

function sinSigno_(arr) { return arr.map(function(v) { return v < 0 ? v + 256 : v; }); }
function sha256_(bytes) { return sinSigno_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)); }
function bytesDe_(t) { return sinSigno_(Utilities.newBlob(t).getBytes()); }
function texto_(bytes) { return Utilities.newBlob(bytes).getDataAsString('UTF-8'); }
function hexDe_(bytes) { return bytes.map(function(v) { return (v < 16 ? '0' : '') + v.toString(16); }).join(''); }
function hex_(h) { return h.match(/../g).map(function(x) { return parseInt(x, 16); }); }
function desdeB64url_(s) {
  s = String(s || '').replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  return sinSigno_(Utilities.base64Decode(s));
}
function b64url_(bytes) {
  return Utilities.base64Encode(bytes.map(function(v) { return v > 127 ? v - 256 : v; })).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

/** Firma de ejemplo para probarMatematica(). */
const PRUEBA = {"x":"b541d0ad16182eb4f5f2c8353ad736e33661de3fae21dd943e619616e56447b6","y":"c1fa2b3134650c46029f0b5f098ecbc944565546ac37bad32ac935c2acb19506","hash":"bcfec93389a7ea9b53cfdca47567966d2edf86e31b6bb6781c957978ac2b480e","r":"1309f32e3e46b420dae2146c67a8f692ebf163f30e0208ba03101ff4556bb81f","s":"5c7264e6c13862af501211c238d22a77910d6d408269d1a2d81aeea181dc7c5d"};
