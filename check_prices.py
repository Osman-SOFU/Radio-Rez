import sqlite3

db = r"C:\Users\osman\Desktop\TEST KLASÖRÜ\data.db"
c = sqlite3.connect(db)

print("DB:", db)
print("TABLES:", [r[0] for r in c.execute("select name from sqlite_master where type='table' order by name")])

q = """
SELECT cp.price_dt, cp.price_odt
FROM channel_prices cp
JOIN channels ch ON ch.id = cp.channel_id
WHERE ch.name = ? AND cp.year = ? AND cp.month = ?
"""
print("SUBAT:", c.execute(q, ("Trt 1 Radyo", 2026, 2)).fetchone())
print("OCAK :", c.execute(q, ("Trt 1 Radyo", 2026, 1)).fetchone())
