import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import random
from collections import defaultdict

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Penjadwalan Tutorial UT",
    page_icon="⚡",
    layout="wide"
)

# =========================================================
# TEMPLATE
# =========================================================
TEMPLATE_COLS = [
    "Kelas/ Lokasi Tutorial",
    "Jum Mhs",
    "SMT",
    "Kode Matakuliah",
    "Mata Kuliah",
    "Nama Tutor",
    "ID Tutor",
    "ID Tutorial",
    "Master Kelas",
    "HARI",
    "JAM",
    "LINK TUWEB (SHORT)"
]

# =========================================================
# SLOT DEFAULT
# =========================================================
DEFAULT_SABTU = """Sabtu 08:00-10:00
Sabtu 10:15-12:15
Sabtu 13:00-15:00
Sabtu 15:15-17:15"""

DEFAULT_MINGGU = """Minggu 08:00-10:00
Minggu 10:15-12:15
Minggu 13:00-15:00
Minggu 15:15-17:15"""

# =========================================================
# NORMALISASI
# =========================================================
def norm_text(x):
    if pd.isna(x):
        return ""
    s = str(x)
    s = re.sub(r"\s+"," ",s).strip()
    return s

def norm_id_digits(x):
    if pd.isna(x):
        return ""
    s = str(x)
    s = s.replace(".0","")
    s = re.sub(r"[^0-9]","",s)
    return s

def tutor_key(row):
    tid = norm_id_digits(row.get("ID Tutor",""))
    if tid!="":
        return f"ID:{tid}"
    return f"NM:{norm_text(row.get('Nama Tutor',''))}"

def kelas_key(row):
    kls = norm_text(row.get("Kelas/ Lokasi Tutorial",""))
    smt = norm_text(row.get("SMT",""))
    return f"{kls} || SMT:{smt}"

# =========================================================
# SLOT PARSER
# =========================================================
def parse_slot_line(line):

    line=line.strip()

    m=re.match(r"^(Sabtu|Minggu)\s+(\d{2}:\d{2}-\d{2}:\d{2})$",line,re.IGNORECASE)

    if not m:
        return None

    return (m.group(1).capitalize(),m.group(2))

def slots_from_multiline(text):

    lines=[x.strip() for x in text.splitlines() if x.strip()]

    slots=[]

    for ln in lines:

        sl=parse_slot_line(ln)

        if sl:
            slots.append(sl)

    return slots

# =========================================================
# PRIORITAS SLOT
# =========================================================
def slot_priority(sl):

    order = {
        ("Minggu","08:00-10:00"):1,
        ("Minggu","10:15-12:15"):2,
        ("Minggu","13:00-15:00"):3,
        ("Minggu","15:15-17:15"):4,

        ("Sabtu","13:00-15:00"):5,
        ("Sabtu","15:15-17:15"):6,

        ("Sabtu","10:15-12:15"):7,
        ("Sabtu","08:00-10:00"):8,
    }

    return order.get(sl,99)

# =========================================================
# TEMPLATE DOWNLOAD
# =========================================================
def make_template():

    df=pd.DataFrame(columns=TEMPLATE_COLS)

    buf=BytesIO()

    with pd.ExcelWriter(buf,engine="openpyxl") as writer:

        df.to_excel(writer,index=False)

    buf.seek(0)

    return buf

# =========================================================
# BUILD DOMAIN SLOT
# =========================================================
def build_tasks(df,slots_sabtu,slots_minggu,tutor_forbidden=None):

    if tutor_forbidden is None:
        tutor_forbidden={}

    mk_count=df.groupby("KELAS_KEY")["Kode Matakuliah"].nunique().to_dict()

    uniq=(df
          .sort_values(["KELAS_KEY","__order__"])
          .drop_duplicates(["KELAS_KEY","Kode Matakuliah"])
          .copy())

    minggu=slots_minggu[:4]

    sabtu_full=slots_sabtu[:4]

    sabtu_normal=[
        ("Sabtu","13:00-15:00"),
        ("Sabtu","15:15-17:15")
    ]

    sabtu_7mk=[
        ("Sabtu","10:15-12:15"),
        ("Sabtu","13:00-15:00"),
        ("Sabtu","15:15-17:15")
    ]

    tasks=[]

    for _,r in uniq.iterrows():

        kkey=r["KELAS_KEY"]

        kode=r["Kode Matakuliah"]

        tutor=r["TutorKey"]

        nmk=int(mk_count.get(kkey,0))

        if nmk in [3,4]:

            domain=minggu

        elif nmk==7:

            domain=sabtu_7mk+minggu

        elif nmk>=8:

            domain=sabtu_full+minggu

        else:

            domain=minggu+sabtu_normal

        forb=tutor_forbidden.get(tutor,set())

        domain=[x for x in domain if x not in forb]

        if len(domain)==0:

            domain=sabtu_full+minggu

        tasks.append({
            "kkey":kkey,
            "kode":kode,
            "tutor":tutor,
            "domain":domain
        })

    return tasks

# =========================================================
# SOLVER
# =========================================================
def greedy_solver(tasks):

    used_class=defaultdict(set)

    used_tutor=defaultdict(set)

    assign={}

    for t in tasks:

        k=t["kkey"]

        kode=t["kode"]

        tutor=t["tutor"]

        for sl in t["domain"]:

            if sl not in used_class[k] and sl not in used_tutor[tutor]:

                assign[(k,kode)]=sl

                used_class[k].add(sl)

                used_tutor[tutor].add(sl)

                break

    return assign

# =========================================================
# SCHEDULER
# =========================================================
def schedule(df,slots_sabtu,slots_minggu,tutor_forbidden):

    tasks=build_tasks(df,slots_sabtu,slots_minggu,tutor_forbidden)

    assign=greedy_solver(tasks)

    map_df=pd.DataFrame([

        {"KELAS_KEY":k[0],"Kode Matakuliah":k[1],"HARI":v[0],"JAM":v[1]}

        for k,v in assign.items()

    ])

    out=df.merge(map_df,on=["KELAS_KEY","Kode Matakuliah"],how="left")

    return out

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:

    st.header("Setting Slot")

    txt_sabtu=st.text_area("Slot Sabtu",DEFAULT_SABTU)

    txt_minggu=st.text_area("Slot Minggu",DEFAULT_MINGGU)

    slots_sabtu=slots_from_multiline(txt_sabtu)

    slots_minggu=slots_from_multiline(txt_minggu)


st.sidebar.caption("© 2026 Haryono – UT Semarang")
# =========================================================
# UI
# =========================================================
st.title("⚡ Penjadwalan Tutorial UT")

st.download_button(
    "Download Template",
    data=make_template(),
    file_name="template_jadwal.xlsx"
)

uploaded=st.file_uploader("Upload Excel",type=["xlsx"])

if not uploaded:
    st.stop()

df=pd.read_excel(uploaded)

df["Kelas/ Lokasi Tutorial"]=df["Kelas/ Lokasi Tutorial"].ffill()
df["SMT"]=df["SMT"].ffill()

df=df[df["Kode Matakuliah"].notna()]

df["TutorKey"]=df.apply(tutor_key,axis=1)

df["KELAS_KEY"]=df.apply(kelas_key,axis=1)

df["__order__"]=np.arange(len(df))

# =========================================================
# PREFERENSI TUTOR
# =========================================================
st.subheader("Preferensi Tutor (Slot TIDAK BOLEH)")

base_unique=df.drop_duplicates(["TutorKey"]).copy()

base_unique["TutorLabel"]=base_unique.apply(
    lambda r:f"{norm_id_digits(r.get('ID Tutor',''))} - {norm_text(r.get('Nama Tutor',''))}".strip(" -"),
    axis=1
)

tutor_options=base_unique.sort_values("TutorLabel")[["TutorKey","TutorLabel"]].values.tolist()

tutor_map={lbl:key for key,lbl in tutor_options}

ALLSLOTS=sorted(slots_sabtu+slots_minggu,key=slot_priority)

slot_labels=[f"{h} | {j}" for (h,j) in ALLSLOTS]

slot_map={f"{h} | {j}":(h,j) for (h,j) in ALLSLOTS}

chosen_tutors=st.multiselect(
    "Pilih tutor yang akan diberi larangan slot",
    options=[lbl for _,lbl in tutor_options]
)

tutor_forbidden={}

for lbl in chosen_tutors:

    tkey=tutor_map[lbl]

    forb=st.multiselect(
        f"Slot tidak boleh: {lbl}",
        options=slot_labels
    )

    tutor_forbidden[tkey]=set(slot_map[x] for x in forb)

# =========================================================
# RUN
# =========================================================
if st.button("Buat Jadwal"):

    jadwal=schedule(df,slots_sabtu,slots_minggu,tutor_forbidden)

    st.success("Jadwal berhasil dibuat")

    st.dataframe(jadwal,use_container_width=True)

    buf=BytesIO()

    with pd.ExcelWriter(buf,engine="openpyxl") as writer:

        jadwal.to_excel(writer,index=False)

    st.download_button(
        "Download Excel Jadwal",
        data=buf.getvalue(),
        file_name="jadwal_ut.xlsx"
    )