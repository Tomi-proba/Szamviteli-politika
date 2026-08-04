#!/usr/bin/env python3
"""
Segédeszköz az iroda_source.html-ben található jelszó-hash cseréjéhez.

FONTOS: ez a jelszó csak VISSZATARTÁS, nem valódi biztonság - az oldal
statikus HTML/JS, nincs szerver, ezért egy technikailag értő személy a
forráskódból (vagy a böngésző DevTools-ából) ki tudná olvasni a hash
kiszámításának módját, és offline brute-force-olhatná. Ez elég egy átlag
ügyfél véletlen hozzáférésének megakadályozására, nem egy valódi
hozzáférés-vezérlés helyettesítője.

Használat:
    python3 webapp-client/hash_password.py "ÚjJelszó"

A kiírt hexa hash-t másold be az iroda_source.html PASSWORD_HASH
konstansába, majd futtasd újra a build.py-t.
"""
import hashlib
import sys

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Használat: python3 hash_password.py <jelszó>")
        sys.exit(1)
    password = sys.argv[1]
    print(hashlib.sha256(password.encode("utf-8")).hexdigest())
