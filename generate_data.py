"""
Dummy data generator for the CRM hierarchy reconciliation assessment (Kenya edition).
All network and facility names are fictional; towns and postal codes are Kenyan.

Creates:
  Partner_Networks.csv    - 10 national healthcare / senior-care partner networks
  Sub_Structures.csv      - 40 county branches, each linked to a parent network
  Alias_Map.csv           - known name variants -> official network name
  CRM_Facilities.csv      - 15,000+ local facility records WITH injected errors
  Injected_Errors_Log.csv - ground truth of every error planted (for the QA check)
"""
import random
import pandas as pd
random.seed(42)
N_ROWS = 15000
networks = [
    ("NET01", "Afya Bora Care Network"),
    ("NET02", "Jamii Health Alliance"),
    ("NET03", "Uzima Senior Care Group"),
    ("NET04", "Tumaini Wellness Federation"),
    ("NET05", "Neema Healthcare Network"),
    ("NET06", "Baraka Elder Care Homes"),
    ("NET07", "Imani Community Health Partners"),
    ("NET08", "Faraja Senior Living Network"),
    ("NET09", "Mwangaza Medical Union"),
    ("NET10", "Pwani Care Alliance"),
]
df_net = pd.DataFrame(networks, columns=["Network_ID", "Network_Name"])
df_net["Country"] = "Kenya"
df_net["Contract_Tier"] = random.choices(["Gold", "Silver", "Bronze"], k=len(df_net))
towns = [  # (postal code, town, county, region)
    ("00100", "Nairobi", "Nairobi", "Nairobi"), ("00800", "Westlands", "Nairobi", "Nairobi"),
    ("00900", "Kiambu", "Kiambu", "Central"), ("01000", "Thika", "Kiambu", "Central"),
    ("10100", "Nyeri", "Nyeri", "Central"), ("10200", "Murang'a", "Murang'a", "Central"),
    ("10300", "Kerugoya", "Kirinyaga", "Central"), ("10400", "Nanyuki", "Laikipia", "Rift Valley"),
    ("20100", "Nakuru", "Nakuru", "Rift Valley"), ("20200", "Kericho", "Kericho", "Rift Valley"),
    ("20300", "Nyahururu", "Laikipia", "Rift Valley"), ("20400", "Bomet", "Bomet", "Rift Valley"),
    ("20500", "Narok", "Narok", "Rift Valley"), ("30100", "Eldoret", "Uasin Gishu", "Rift Valley"),
    ("30200", "Kitale", "Trans Nzoia", "Rift Valley"), ("30300", "Kapsabet", "Nandi", "Rift Valley"),
    ("30500", "Lodwar", "Turkana", "Rift Valley"), ("30700", "Iten", "Elgeyo Marakwet", "Rift Valley"),
    ("40100", "Kisumu", "Kisumu", "Nyanza"), ("40200", "Kisii", "Kisii", "Nyanza"),
    ("40300", "Homa Bay", "Homa Bay", "Nyanza"), ("40400", "Migori", "Migori", "Nyanza"),
    ("40500", "Nyamira", "Nyamira", "Nyanza"), ("40600", "Siaya", "Siaya", "Nyanza"),
    ("50100", "Kakamega", "Kakamega", "Western"), ("50200", "Bungoma", "Bungoma", "Western"),
    ("50300", "Vihiga", "Vihiga", "Western"), ("50400", "Busia", "Busia", "Western"),
    ("60100", "Embu", "Embu", "Eastern"), ("60200", "Meru", "Meru", "Eastern"),
    ("60400", "Chuka", "Tharaka Nithi", "Eastern"), ("60500", "Marsabit", "Marsabit", "Eastern"),
    ("70100", "Garissa", "Garissa", "North Eastern"), ("70200", "Wajir", "Wajir", "North Eastern"),
    ("80100", "Mombasa", "Mombasa", "Coast"), ("80200", "Malindi", "Kilifi", "Coast"),
    ("80300", "Voi", "Taita Taveta", "Coast"), ("80400", "Ukunda", "Kwale", "Coast"),
    ("90100", "Machakos", "Machakos", "Eastern"), ("90200", "Kitui", "Kitui", "Eastern"),
]
random.shuffle(towns)
subs = []
for i, (pc, town, county, region) in enumerate(towns):
    net_id, net_name = networks[i // 4]
    subs.append({
        "SubStructure_ID": f"SUB{i+1:03d}",
        "SubStructure_Name": f"{net_name.split()[0]} {town} Branch",
        "Parent_Network_ID": net_id,
        "Postal_Code": pc,
        "Town": town,
        "County": county,
        "Region": region,
    })
df_sub = pd.DataFrame(subs)
aliases = [
    ("Afya Bora", "NET01"), ("ABCN", "NET01"), ("Afya Bora Network", "NET01"),
    ("Jamii Health", "NET02"), ("JHA", "NET02"),
    ("Uzima Seniors", "NET03"), ("Uzima Care", "NET03"),
    ("Tumaini Wellness", "NET04"), ("TWF", "NET04"),
    ("Neema Health", "NET05"), ("Neema Network", "NET05"),
    ("Baraka Homes", "NET06"), ("Baraka Elder Care", "NET06"),
    ("Imani Health", "NET07"), ("ICHP", "NET07"),
    ("Faraja Seniors", "NET08"), ("Faraja Living", "NET08"),
    ("Mwangaza Medical", "NET09"), ("MMU", "NET09"),
    ("Pwani Care", "NET10"), ("Pwani Alliance", "NET10"),
]
df_alias = pd.DataFrame(aliases, columns=["Alias", "Network_ID"])
df_alias = df_alias.merge(df_net[["Network_ID", "Network_Name"]], on="Network_ID")
types = ["Nursing Home", "Health Centre", "Dispensary", "Medical Centre",
         "Elderly Care Home", "Hospice", "Community Clinic", "Rehabilitation Centre"]
names = ["Amani", "Upendo", "Baraka", "Tumaini", "Neema", "St. Mary's", "St. Joseph's",
         "Mlimani", "Riverside", "Hekima", "Rehema", "Nuru", "Furaha", "Matumaini",
         "Shalom", "Bethany", "Mercy", "Kijani", "Savanna", "Acacia", "Jacaranda",
         "Mwanzo", "Pendo", "Bidii", "Umoja"]
areas = ["Central", "East", "West", "North", "South", "Town", "Estate", "Road",
         "Junction", "Hill", "Gardens", "Market"]
statuses = ["Active", "Pending", "Expired", "Terminated"]
status_w = [0.6, 0.15, 0.15, 0.10]
net_name_by_id = dict(networks)
rows, used_names = [], set()
for i in range(N_ROWS):
    s = random.choice(subs)
    while True: 
        nm = f"{random.choice(names)} {random.choice(types)} - {s['Town']} {random.choice(areas)}"
        if nm not in used_names:
            used_names.add(nm)
            break
        nm = None
    rows.append({
        "Facility_ID": f"FAC{i+1:06d}",
        "Facility_Name": nm,
        "Postal_Code": s["Postal_Code"],
        "City": s["Town"],
        "Agreement_Status": random.choices(statuses, status_w)[0],
        "Parent_Network": net_name_by_id[s["Parent_Network_ID"]],
        "Sub_Structure": s["SubStructure_Name"],
        "Last_Updated": pd.Timestamp("2025-01-01") + pd.Timedelta(days=random.randint(0, 600)),
    })
df = pd.DataFrame(rows)
log = []
idx = list(df.index)
random.shuffle(idx)
pos = 0
def take(pct):
    global pos
    n = int(N_ROWS * pct)
    chosen = idx[pos:pos + n]
    pos += n
    return chosen
def typo(original):
    """Return a corrupted version of the name that is always different."""
    for _ in range(20):
        s = list(original)
        j = random.randint(1, len(s) - 2)
        op = random.choice(["swap", "drop", "case", "space"])
        if op == "swap":
            s[j], s[j + 1] = s[j + 1], s[j]
            out = "".join(s)
        elif op == "drop":
            del s[j]
            out = "".join(s)
        elif op == "case":
            out = "".join(s).upper()
        else:
            out = "  " + "".join(s).lower() + " "
        if out != original:
            return out
    return original[:-1]
for r in take(0.05):
    log.append((df.at[r, "Facility_ID"], "Missing parent", df.at[r, "Parent_Network"]))
    df.at[r, "Parent_Network"] = ""
for r in take(0.04):
    orig = df.at[r, "Parent_Network"]
    nid = [k for k, v in net_name_by_id.items() if v == orig][0]
    if random.random() < 0.5:
        new = random.choice([a for a, n in aliases if n == nid])
        kind = "Alias / abbreviation used"
    else:
        new = typo(orig)
        kind = "Typo in parent name"
    df.at[r, "Parent_Network"] = new
    log.append((df.at[r, "Facility_ID"], kind, orig))
for r in take(0.03):
    orig = df.at[r, "Parent_Network"]
    wrong = random.choice([n for _, n in networks if n != orig])
    df.at[r, "Parent_Network"] = wrong
    log.append((df.at[r, "Facility_ID"], "Parent conflicts with sub-structure", orig))
for r in take(0.02):
    orig = df.at[r, "Sub_Structure"]
    df.at[r, "Sub_Structure"] = random.choice(
        ["Unknown Regional Office", "Legacy Group West", "Independent", "TBD", "County Office"])
    log.append((df.at[r, "Facility_ID"], "Unmapped sub-structure", orig))
for r in take(0.02):
    orig = df.at[r, "Parent_Network"]
    df.at[r, "Parent_Network"] = df.at[r, "Sub_Structure"]
    log.append((df.at[r, "Facility_ID"], "Sub-structure entered as parent", orig))
for r in take(0.015):
    df.at[r, "Postal_Code"] = ""
    log.append((df.at[r, "Facility_ID"], "Missing postal code", ""))
for r in take(0.015):
    df.at[r, "Agreement_Status"] = ""
    log.append((df.at[r, "Facility_ID"], "Missing agreement status", ""))
for r in take(0.01):
    df.at[r, "Postal_Code"] = random.choice(["ABC12", "123", "0000000", "N/A"])
    log.append((df.at[r, "Facility_ID"], "Invalid postal code", ""))
for r in take(0.01):
    df.at[r, "Parent_Network"] = ""
    df.at[r, "Sub_Structure"] = ""
    log.append((df.at[r, "Facility_ID"], "Missing parent and sub-structure", ""))
for r in take(0.004):
    df.at[r, "Parent_Network"] = ""
    df.at[r, "Sub_Structure"] = ""
    df.at[r, "Postal_Code"] = random.choice(["", "N/A", "UNKNOWN", "99999"])
    log.append((df.at[r, "Facility_ID"], "Fully orphaned record", ""))
dups = df.sample(150, random_state=1).copy()
dups["Facility_ID"] = [f"FAC{N_ROWS + k + 1:06d}" for k in range(len(dups))]
for fid in dups["Facility_ID"]:
    log.append((fid, "Duplicate facility record", ""))
df = pd.concat([df, dups]).sample(frac=1, random_state=7).reset_index(drop=True)
df["Last_Updated"] = df["Last_Updated"].dt.strftime("%Y-%m-%d")
df.to_csv("CRM_Facilities.csv", index=False, encoding="utf-8-sig")
df_net.to_csv("Partner_Networks.csv", index=False, encoding="utf-8-sig")
df_sub.to_csv("Sub_Structures.csv", index=False, encoding="utf-8-sig")
df_alias.to_csv("Alias_Map.csv", index=False, encoding="utf-8-sig")
pd.DataFrame(log, columns=["Facility_ID", "Injected_Error", "True_Value"]).to_csv(
    "Injected_Errors_Log.csv", index=False, encoding="utf-8-sig")

print(f"Facilities: {len(df)} rows | errors injected: {len(log)}")
