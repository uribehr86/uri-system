"""
Generate self-signed certificate valid for both 127.0.0.1 and 192.168.1.37
"""
import datetime, ipaddress
from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa

# Generate key
key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

# Subject
subject = issuer = x509.Name([
    x509.NameAttribute(NameOID.COMMON_NAME, u"URI System"),
])

# Certificate
cert = (
    x509.CertificateBuilder()
    .subject_name(subject)
    .issuer_name(issuer)
    .public_key(key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(datetime.datetime.utcnow())
    .not_valid_after(datetime.datetime.utcnow() + datetime.timedelta(days=3650))
    .add_extension(
        x509.SubjectAlternativeName([
            x509.DNSName(u"localhost"),
            x509.IPAddress(ipaddress.IPv4Address(u"127.0.0.1")),
            x509.IPAddress(ipaddress.IPv4Address(u"192.168.1.10")),
            x509.IPAddress(ipaddress.IPv4Address(u"192.168.1.37")),
            x509.IPAddress(ipaddress.IPv4Address(u"10.0.0.31")),
        ]),
        critical=False,
    )
    .sign(key, hashes.SHA256())
)

# Save cert
with open("server.crt", "wb") as f:
    f.write(cert.public_bytes(serialization.Encoding.PEM))

# Save key
with open("server.key", "wb") as f:
    f.write(key.private_bytes(
        serialization.Encoding.PEM,
        serialization.PrivateFormat.TraditionalOpenSSL,
        serialization.NoEncryption()
    ))

print("Done! server.crt and server.key created")
print("Valid for: 127.0.0.1, 192.168.1.37, 10.0.0.31, localhost")
