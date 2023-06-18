from archicad import ACConnection

conn = ACConnection.connect()
assert conn, "Nie otworzono żadnego projektu Archicada."

acc = conn.commands
act = conn.types
acu = conn.utilities