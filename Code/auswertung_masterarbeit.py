# ============================================================
# AUSWERTUNG DES ELAN-ANNOTATIONSKORPUS
# Masterarbeit: Propagandastrategien in der Deutschen Wochenschau
# ============================================================
#
# Das Skript:
# 1. liest alle .eaf-Dateien eines Ordners ein
# 2. verbindet Segment mit Content, Actor, Topic,
#    Propaganda_Strategies, Visual_Technique und Audio
# 3. berechnet Start, Ende und Dauer jedes Segments
# 4. trennt Mehrfachannotationen
# 5. erstellt statistische Auswertungen
# 6. exportiert CSV-Dateien
# 7. erstellt eine Excel-Gesamtauswertung
# 8. erzeugt PNG-Grafiken (300 dpi)
#
# Benötigte Pakete:
# pip install pandas matplotlib openpyxl
#
# ============================================================

from pathlib import Path
import re
import xml.etree.ElementTree as ET
from collections import Counter

import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# 1. EINSTELLUNGEN
# ============================================================

# Portabler Pfad: Der Ordner mit den EAF-Dateien liegt neben diesem Skript.
# Dadurch enthält das Skript keinen persönlichen oder systemspezifischen Windows-Pfad.
SKRIPT_ORDNER = Path(__file__).resolve().parent
EAF_ORDNER = SKRIPT_ORDNER / "Annotationskorpus EAF"

# Ausgabeordner wird automatisch neben dem Skript erstellt.
AUSGABE_ORDNER = SKRIPT_ORDNER / "Auswertung"

DATEN_ORDNER = AUSGABE_ORDNER / "Daten"
ABBILDUNGEN_ORDNER = AUSGABE_ORDNER / "Abbildungen"
ERGEBNISSE_ORDNER = AUSGABE_ORDNER / "Ergebnisse"

for ordner in [
    AUSGABE_ORDNER,
    DATEN_ORDNER,
    ABBILDUNGEN_ORDNER,
    ERGEBNISSE_ORDNER,
]:
    ordner.mkdir(parents=True, exist_ok=True)


# ============================================================
# 2. ZUORDNUNG FOLGE -> JAHR
# ============================================================

FOLGE_JAHR = {
    # 1940
    511: 1940,
    520: 1940,
    533: 1940,

    # 1941
    559: 1941,
    568: 1941,
    572: 1941,
    585: 1941,

    # 1942
    598: 1942,
    611: 1942,
    621: 1942,
    637: 1942,

    # 1943
    650: 1943,
    656: 1943,
    663: 1943,
    676: 1943,
    681: 1943,
    689: 1943,

    # 1944
    702: 1944,
    715: 1944,
    728: 1944,
    734: 1944,
    741: 1944,

    # 1945
    748: 1945,
    750: 1945,
    755: 1945,
}


# ============================================================
# 3. TIER-NAMEN
# ============================================================

TIER_SEGMENT = "Segment"
TIER_CONTENT = "Content"
TIER_ACTOR = "Actor"
TIER_TOPIC = "Topic"
TIER_STRATEGY = "Propaganda_Strategies"
TIER_VISUAL = "Visual_Technique"
TIER_AUDIO = "Audio"
TIER_NOTES = "Notes"


# ============================================================
# 4. OPTIONALE NORMALISIERUNG
# ============================================================
#
# Hier können Schreibvarianten zusammengeführt werden.
# Deine EAF-Dateien werden dadurch NICHT verändert.
#
# WICHTIG:
# Prüfe diese Zuordnungen vor der endgültigen Auswertung.
# Nur Varianten aufnehmen, die tatsächlich dasselbe bedeuten.
# ============================================================

NORMALISIERUNG = {

    # Actor: Abkürzungen und ausgeschriebene gleichbedeutende Varianten
    "HJ": "Hitlerjugend",
    "Hitlerjugend": "Hitlerjugend",
    "BDM": "Bund Deutscher Mädel",
    "Bund Deutscher Mädel": "Bund Deutscher Mädel",
    "RAD": "Reichsarbeitsdienst",
    "Reichsarbeitsdienst": "Reichsarbeitsdienst",
    "SA": "SA",
    "Sturmabteilung": "SA",
    "SS": "SS",
    "Schutzstaffel": "SS",


    # Propagandastrategien
    # Die Bezeichnungen entsprechen der finalen Annotationsguideline.
    "Personenkult": "Personenkult",
    "personenkult": "Personenkult",
    "Heroisierung": "Heroisierung",
    "heroisierung": "Heroisierung",
    "Feindbildkonstruktion": "Feindbildkonstruktion",
    "feindbildkonstruktion": "Feindbildkonstruktion",
    "Volksgemeinschaft": "Volksgemeinschaft",
    "volksgemeinschaft": "Volksgemeinschaft",
    "Opferbereitschaft und Durchhalten": "Opferbereitschaft und Durchhalten",
    "opferbereitschaft und durchhalten": "Opferbereitschaft und Durchhalten",
    "Legitimierung": "Legitimierung",
    "legitimierung": "Legitimierung",
    "Dämonisierung": "Dämonisierung",
    "dämonisierung": "Dämonisierung",
    "Rassenideologie": "Rassenideologie",
    "rassenideologie": "Rassenideologie",

    # Audio – reine Groß-/Kleinschreibungsvarianten
    "wertender Sprecher": "Wertender Sprecher",
    "Wertender Sprecher": "Wertender Sprecher",
    "originalton": "Originalton",
    "Originalton": "Originalton",
    "publikumsreaktion": "Publikumsreaktion",
    "Publikumsreaktion": "Publikumsreaktion",
    "musik": "Musik",
    "Musik": "Musik",

    # Visual – reine Groß-/Kleinschreibungsvarianten
    "nahaufnahme": "Nahaufnahme",
    "Nahaufnahme": "Nahaufnahme",
    "totale": "Totale",
    "Totale": "Totale",
    "massendarstellung": "Massendarstellung",
    "Massendarstellung": "Massendarstellung",
    "symbolik": "Symbolik",
    "Symbolik": "Symbolik",
    "kamerabewegung": "Kamerabewegung",
    "Kamerabewegung": "Kamerabewegung",
    "kamerafahrt": "Kamerafahrt",
    "Kamerafahrt": "Kamerafahrt",
    "luftaufnahme": "Luftaufnahme",
    "Luftaufnahme": "Luftaufnahme",
    "animation": "Animation",
    "Animation": "Animation",
}


# ============================================================
# 5. HILFSFUNKTIONEN
# ============================================================

def clean_text(text):
    """Bereinigt Annotationstext."""
    if text is None:
        return ""

    return " ".join(text.split()).strip()


def normalize_value(value):
    """
    Normalisiert bekannte Schreibvarianten ohne Beachtung der
    Groß-/Kleinschreibung.

    Dadurch werden reine Schreibvarianten wie
    "italienische Soldaten" und "Italienische Soldaten"
    als dieselbe Kategorie behandelt. Die Schreibweise des
    kanonischen Werts aus NORMALISIERUNG wird beibehalten.
    """
    value = clean_text(value)

    if not value:
        return ""

    # Explizite Normalisierungen case-insensitiv anwenden.
    normalisierung_casefold = {
        clean_text(key).casefold(): target
        for key, target in NORMALISIERUNG.items()
    }

    return normalisierung_casefold.get(value.casefold(), value)


def split_values(value):
    """
    Trennt Mehrfachannotationen.

    In den geprüften EAF-Dateien werden mehrere Kategorien
    überwiegend durch Kommas getrennt.

    Semikolon wird vorsichtshalber ebenfalls unterstützt.
    """
    value = clean_text(value)

    if not value:
        return []

    parts = re.split(r"\s*[,;]\s*", value)

    cleaned = []

    for part in parts:
        part = normalize_value(part)

        if part and part not in cleaned:
            cleaned.append(part)

    return cleaned


def list_to_string(values):
    """Liste wieder als lesbaren String speichern."""
    if not values:
        return ""

    return "; ".join(values)


def extract_episode_number(filename):
    """
    Extrahiert die Wochenschau-Folgennummer aus dem Dateinamen.

    Beispiele:
    DW_572.eaf
    DDW_572_BA.eaf
    35623_1_1_DDW_572_BA.eaf
    """
    name = Path(filename).stem

    # bevorzugt Nummer nach DW / DDW
    match = re.search(r"(?:DDW|DW)[_\-\s]*(\d{3})", name, re.IGNORECASE)

    if match:
        return int(match.group(1))

    # Fallback: bekannte Folgennummer im Dateinamen suchen
    numbers = re.findall(r"\d{3}", name)

    for number in numbers:
        number = int(number)

        if number in FOLGE_JAHR:
            return number

    return None


# ============================================================
# 6. EAF-DATEI EINLESEN
# ============================================================

def parse_eaf(file_path):

    tree = ET.parse(file_path)
    root = tree.getroot()

    # --------------------------------------------------------
    # TIME SLOTS
    # --------------------------------------------------------

    time_slots = {}

    time_order = root.find("TIME_ORDER")

    if time_order is not None:

        for slot in time_order.findall("TIME_SLOT"):

            slot_id = slot.get("TIME_SLOT_ID")
            time_value = slot.get("TIME_VALUE")

            if time_value is not None:
                time_slots[slot_id] = int(time_value)

    # --------------------------------------------------------
    # Alle Annotationen sammeln
    # --------------------------------------------------------

    annotations = {}

    tier_annotations = {}

    for tier in root.findall("TIER"):

        tier_id = tier.get("TIER_ID")

        tier_annotations[tier_id] = []

        for annotation_wrapper in tier.findall("ANNOTATION"):

            annotation = None
            annotation_type = None

            alignable = annotation_wrapper.find("ALIGNABLE_ANNOTATION")
            ref_annotation = annotation_wrapper.find("REF_ANNOTATION")

            if alignable is not None:
                annotation = alignable
                annotation_type = "alignable"

            elif ref_annotation is not None:
                annotation = ref_annotation
                annotation_type = "ref"

            if annotation is None:
                continue

            annotation_id = annotation.get("ANNOTATION_ID")

            value_element = annotation.find("ANNOTATION_VALUE")

            value = ""

            if value_element is not None:
                value = clean_text(value_element.text)

            info = {
                "id": annotation_id,
                "tier": tier_id,
                "type": annotation_type,
                "value": value,
                "ref": annotation.get("ANNOTATION_REF"),
                "time1": annotation.get("TIME_SLOT_REF1"),
                "time2": annotation.get("TIME_SLOT_REF2"),
            }

            annotations[annotation_id] = info
            tier_annotations[tier_id].append(info)

    # --------------------------------------------------------
    # Segment-Annotationen
    # --------------------------------------------------------

    segments = tier_annotations.get(TIER_SEGMENT, [])

    rows = []

    folge = extract_episode_number(file_path.name)

    jahr = FOLGE_JAHR.get(folge)

    if folge is None:
        print(
            f"WARNUNG: Folgennummer konnte nicht erkannt werden: "
            f"{file_path.name}"
        )

    elif jahr is None:
        print(
            f"WARNUNG: Für Folge {folge} ist kein Jahr hinterlegt."
        )

    # --------------------------------------------------------
    # Direkte Referenzen auf Segment sammeln
    # --------------------------------------------------------

    refs_by_parent = {}

    for tier_id, anns in tier_annotations.items():

        if tier_id == TIER_SEGMENT:
            continue

        for ann in anns:

            parent = ann["ref"]

            if parent:
                refs_by_parent.setdefault(parent, {}).setdefault(
                    tier_id, []
                ).append(ann["value"])

    # --------------------------------------------------------
    # Segmente zusammensetzen
    # --------------------------------------------------------

    for segment_index, segment in enumerate(segments, start=1):

        segment_id = segment["id"]

        start_ms = time_slots.get(segment["time1"])
        end_ms = time_slots.get(segment["time2"])

        if start_ms is not None and end_ms is not None:
            duration_ms = end_ms - start_ms
        else:
            duration_ms = None

        dependent = refs_by_parent.get(segment_id, {})

        def get_tier_value(tier_name):

            values = dependent.get(tier_name, [])

            values = [
                clean_text(v)
                for v in values
                if clean_text(v)
            ]

            return "; ".join(values)

        content = get_tier_value(TIER_CONTENT)
        actor_raw = get_tier_value(TIER_ACTOR)
        topic_raw = get_tier_value(TIER_TOPIC)
        strategy_raw = get_tier_value(TIER_STRATEGY)
        visual_raw = get_tier_value(TIER_VISUAL)
        audio_raw = get_tier_value(TIER_AUDIO)
        notes = get_tier_value(TIER_NOTES)

        actors = split_values(actor_raw)
        topics = split_values(topic_raw)
        strategies = split_values(strategy_raw)
        visuals = split_values(visual_raw)
        audios = split_values(audio_raw)

        rows.append({
            "Datei": file_path.name,
            "Folge": folge,
            "Jahr": jahr,
            "Segment_Nr": segment_index,
            "Segment_ID": segment_id,

            "Start_ms": start_ms,
            "Ende_ms": end_ms,

            "Start_Sekunden":
                start_ms / 1000 if start_ms is not None else None,

            "Ende_Sekunden":
                end_ms / 1000 if end_ms is not None else None,

            "Dauer_Sekunden":
                duration_ms / 1000
                if duration_ms is not None else None,

            "Dauer_Minuten":
                duration_ms / 60000
                if duration_ms is not None else None,

            "Content": content,

            "Actor": list_to_string(actors),
            "Topic": list_to_string(topics),
            "Propaganda_Strategies": list_to_string(strategies),
            "Visual_Technique": list_to_string(visuals),
            "Audio": list_to_string(audios),

            "Notes": notes,

            # interne Listen für spätere Analyse
            "_actors": actors,
            "_topics": topics,
            "_strategies": strategies,
            "_visuals": visuals,
            "_audios": audios,
        })

    return rows


# ============================================================
# 7. ALLE 25 EAF-DATEIEN EINLESEN
# ============================================================

eaf_files = sorted(EAF_ORDNER.glob("*.eaf"))

print()
print("=" * 70)
print("EAF-AUSWERTUNG")
print("=" * 70)
print(f"Gefundene EAF-Dateien: {len(eaf_files)}")

if len(eaf_files) == 0:
    raise FileNotFoundError(
        f"Keine EAF-Dateien gefunden in:\n{EAF_ORDNER}"
    )

if len(eaf_files) != 25:
    print(
        f"ACHTUNG: Erwartet werden 25 Dateien, "
        f"gefunden wurden {len(eaf_files)}."
    )

all_rows = []

for file in eaf_files:

    print(f"Lese: {file.name}")

    try:
        rows = parse_eaf(file)
        all_rows.extend(rows)

    except Exception as error:
        print(f"FEHLER bei {file.name}: {error}")


df = pd.DataFrame(all_rows)

if df.empty:
    raise ValueError("Es konnten keine Segmente eingelesen werden.")


# ============================================================
# 7a. GROSS-/KLEINSCHREIBUNG KORPUSWEIT VEREINHEITLICHEN
# ============================================================
#
# Werte, die sich ausschließlich durch Groß-/Kleinschreibung
# unterscheiden, werden als dieselbe Kategorie behandelt.
# Inhaltlich unterschiedliche Bezeichnungen werden NICHT
# automatisch zusammengeführt.
#
# Beispiel:
# "italienische Soldaten" + "Italienische Soldaten"
# -> eine gemeinsame Kategorie.
# ============================================================

def canonicalize_case_variants(series_of_lists):
    """Vereinheitlicht ausschließlich Groß-/Kleinschreibungsvarianten."""
    variants = {}

    for values in series_of_lists:
        for value in values:
            key = value.casefold()
            variants.setdefault(key, [])
            if value not in variants[key]:
                variants[key].append(value)

    canonical = {}

    for key, spellings in variants.items():
        # Bevorzugt eine Variante, die mit Großbuchstaben beginnt.
        preferred = next(
            (v for v in spellings if v and v[0].isupper()),
            spellings[0]
        )
        canonical[key] = preferred

    def convert(values):
        result = []
        for value in values:
            normalized = canonical.get(value.casefold(), value)
            if normalized not in result:
                result.append(normalized)
        return result

    return series_of_lists.apply(convert)


for list_column in [
    "_actors",
    "_topics",
    "_strategies",
    "_visuals",
    "_audios",
]:
    df[list_column] = canonicalize_case_variants(df[list_column])

# Öffentliche String-Spalten nach der Vereinheitlichung aktualisieren.
df["Actor"] = df["_actors"].apply(list_to_string)
df["Topic"] = df["_topics"].apply(list_to_string)
df["Propaganda_Strategies"] = df["_strategies"].apply(list_to_string)
df["Visual_Technique"] = df["_visuals"].apply(list_to_string)
df["Audio"] = df["_audios"].apply(list_to_string)


# ============================================================
# 8. SORTIEREN
# ============================================================

df = df.sort_values(
    ["Jahr", "Folge", "Segment_Nr"],
    na_position="last"
).reset_index(drop=True)


# ============================================================
# 9. ÖFFENTLICHE MASTER-TABELLE
# ============================================================

public_columns = [
    "Datei",
    "Folge",
    "Jahr",
    "Segment_Nr",
    "Segment_ID",
    "Start_Sekunden",
    "Ende_Sekunden",
    "Dauer_Sekunden",
    "Dauer_Minuten",
    "Content",
    "Actor",
    "Topic",
    "Propaganda_Strategies",
    "Visual_Technique",
    "Audio",
    "Notes",
]

df_public = df[public_columns].copy()

df_public.to_csv(
    DATEN_ORDNER / "01_alle_segmente.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 10. LONG-FORMAT FÜR STRATEGIEN
# ============================================================

strategy_rows = []

for _, row in df.iterrows():

    for strategy in row["_strategies"]:

        strategy_rows.append({
            "Datei": row["Datei"],
            "Folge": row["Folge"],
            "Jahr": row["Jahr"],
            "Segment_Nr": row["Segment_Nr"],
            "Segment_ID": row["Segment_ID"],
            "Strategie": strategy,
            "Dauer_Sekunden": row["Dauer_Sekunden"],
            "Dauer_Minuten": row["Dauer_Minuten"],
        })

strategy_long = pd.DataFrame(strategy_rows)


# ============================================================
# 11. LONG-FORMAT-HILFSFUNKTION
# ============================================================

def create_long_dataframe(df_source, list_column, value_name):

    rows = []

    for _, row in df_source.iterrows():

        for value in row[list_column]:

            rows.append({
                "Folge": row["Folge"],
                "Jahr": row["Jahr"],
                "Segment_Nr": row["Segment_Nr"],
                "Segment_ID": row["Segment_ID"],
                value_name: value,
                "Dauer_Sekunden": row["Dauer_Sekunden"],
                "Dauer_Minuten": row["Dauer_Minuten"],
            })

    return pd.DataFrame(rows)


actor_long = create_long_dataframe(
    df, "_actors", "Actor"
)

topic_long = create_long_dataframe(
    df, "_topics", "Topic"
)

visual_long = create_long_dataframe(
    df, "_visuals", "Visual"
)

audio_long = create_long_dataframe(
    df, "_audios", "Audio"
)


# ============================================================
# 12. KORPUSSTATISTIK
# ============================================================

korpusstatistik = (
    df.groupby("Jahr", dropna=False)
    .agg(
        Folgen=("Folge", "nunique"),
        Segmente=("Segment_ID", "count"),
        Gesamtdauer_Sekunden=("Dauer_Sekunden", "sum"),
        Durchschnittliche_Segmentdauer_Sekunden=(
            "Dauer_Sekunden", "mean"
        ),
        Median_Segmentdauer_Sekunden=(
            "Dauer_Sekunden", "median"
        ),
    )
    .reset_index()
)

korpusstatistik["Gesamtdauer_Minuten"] = (
    korpusstatistik["Gesamtdauer_Sekunden"] / 60
)

korpusstatistik.to_csv(
    DATEN_ORDNER / "02_korpusstatistik.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 13. STRATEGIEN GESAMT
# ============================================================

if not strategy_long.empty:

    strategien_gesamt = (
        strategy_long.groupby("Strategie")
        .agg(
            Anzahl_Segmente=("Segment_ID", "count"),
            Gesamtdauer_Sekunden=("Dauer_Sekunden", "sum"),
            Gesamtdauer_Minuten=("Dauer_Minuten", "sum"),
        )
        .reset_index()
        .sort_values(
            "Anzahl_Segmente",
            ascending=False
        )
    )

else:
    strategien_gesamt = pd.DataFrame(
        columns=[
            "Strategie",
            "Anzahl_Segmente",
            "Gesamtdauer_Sekunden",
            "Gesamtdauer_Minuten",
        ]
    )

strategien_gesamt.to_csv(
    DATEN_ORDNER / "03_strategien_gesamt.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 14. STRATEGIEN PRO JAHR – ABSOLUT
# ============================================================

if not strategy_long.empty:

    strategien_jahr = (
        strategy_long.groupby(
            ["Jahr", "Strategie"]
        )
        .agg(
            Anzahl_Segmente=("Segment_ID", "count"),
            Gesamtdauer_Sekunden=("Dauer_Sekunden", "sum"),
            Gesamtdauer_Minuten=("Dauer_Minuten", "sum"),
        )
        .reset_index()
    )

else:
    strategien_jahr = pd.DataFrame()


strategien_jahr.to_csv(
    DATEN_ORDNER / "04_strategien_pro_jahr_absolut.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 15. RELATIVE STRATEGIEHÄUFIGKEIT PRO JAHR
# ============================================================
#
# Anteil der Segmente eines Jahres, in denen eine Strategie
# vorkommt.
#
# WICHTIG:
# Da mehrere Strategien pro Segment vorkommen können,
# können die Prozentwerte zusammen > 100 % ergeben.
# ============================================================

segments_per_year = (
    df.groupby("Jahr")["Segment_ID"]
    .count()
    .rename("Segmente_Jahr")
)

if not strategy_long.empty:

    strategie_relativ = (
        strategy_long.groupby(
            ["Jahr", "Strategie"]
        )["Segment_ID"]
        .count()
        .rename("Strategie_Segmente")
        .reset_index()
    )

    strategie_relativ = strategie_relativ.merge(
        segments_per_year,
        on="Jahr",
        how="left"
    )

    strategie_relativ["Anteil_Segmente_Prozent"] = (
        strategie_relativ["Strategie_Segmente"]
        / strategie_relativ["Segmente_Jahr"]
        * 100
    )

else:
    strategie_relativ = pd.DataFrame()


strategie_relativ.to_csv(
    DATEN_ORDNER / "05_strategien_pro_jahr_relativ.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 16. STRATEGIE-LAUFZEIT PRO JAHR
# ============================================================
#
# ACHTUNG:
# Bei Mehrfachannotation wird die Dauer eines Segments jeder
# darin vorkommenden Strategie zugerechnet.
#
# Daher können auch Laufzeitanteile zusammen >100 % ergeben.
# ============================================================

runtime_year = (
    df.groupby("Jahr")["Dauer_Sekunden"]
    .sum()
    .rename("Gesamtdauer_Jahr_Sekunden")
)

if not strategy_long.empty:

    strategie_runtime = (
        strategy_long.groupby(
            ["Jahr", "Strategie"]
        )["Dauer_Sekunden"]
        .sum()
        .rename("Strategie_Dauer_Sekunden")
        .reset_index()
    )

    strategie_runtime = strategie_runtime.merge(
        runtime_year,
        on="Jahr",
        how="left"
    )

    strategie_runtime["Anteil_Laufzeit_Prozent"] = (
        strategie_runtime["Strategie_Dauer_Sekunden"]
        / strategie_runtime["Gesamtdauer_Jahr_Sekunden"]
        * 100
    )

else:
    strategie_runtime = pd.DataFrame()


strategie_runtime.to_csv(
    DATEN_ORDNER / "06_strategien_laufzeit_pro_jahr.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 17. SEGMENTLÄNGEN PRO STRATEGIE
# ============================================================

if not strategy_long.empty:

    strategie_segmentlaenge = (
        strategy_long.groupby("Strategie")
        .agg(
            N=("Dauer_Sekunden", "count"),
            Gesamtdauer_Sekunden=("Dauer_Sekunden", "sum"),
            Gesamtdauer_Minuten=("Dauer_Minuten", "sum"),
            Mittelwert_Sekunden=("Dauer_Sekunden", "mean"),
            Median_Sekunden=("Dauer_Sekunden", "median"),
            Standardabweichung_Sekunden=("Dauer_Sekunden", "std"),
            Minimum_Sekunden=("Dauer_Sekunden", "min"),
            Maximum_Sekunden=("Dauer_Sekunden", "max"),
        )
        .reset_index()
        .sort_values(
            "Median_Sekunden",
            ascending=False
        )
    )

else:
    strategie_segmentlaenge = pd.DataFrame()


strategie_segmentlaenge.to_csv(
    DATEN_ORDNER / "07_segmentlaengen_pro_strategie.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 18. SEGMENTLÄNGEN STRATEGIE × JAHR
# ============================================================

if not strategy_long.empty:

    strategie_segmentlaenge_jahr = (
        strategy_long.groupby(
            ["Jahr", "Strategie"]
        )
        .agg(
            N=("Dauer_Sekunden", "count"),
            Mittelwert_Sekunden=("Dauer_Sekunden", "mean"),
            Median_Sekunden=("Dauer_Sekunden", "median"),
            Standardabweichung_Sekunden=("Dauer_Sekunden", "std"),
            Minimum_Sekunden=("Dauer_Sekunden", "min"),
            Maximum_Sekunden=("Dauer_Sekunden", "max"),
        )
        .reset_index()
    )

else:
    strategie_segmentlaenge_jahr = pd.DataFrame()


strategie_segmentlaenge_jahr.to_csv(
    DATEN_ORDNER / "08_segmentlaengen_strategie_jahr.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 19. STRATEGIE × TOPIC / ACTOR / VISUAL / AUDIO
# ============================================================

def make_cooccurrence(df_source, second_list_column, second_name):

    rows = []

    for _, row in df_source.iterrows():

        strategies = row["_strategies"]
        second_values = row[second_list_column]

        for strategy in strategies:

            for second in second_values:

                rows.append({
                    "Strategie": strategy,
                    second_name: second,
                    "Jahr": row["Jahr"],
                    "Folge": row["Folge"],
                    "Segment_ID": row["Segment_ID"],
                })

    temp = pd.DataFrame(rows)

    if temp.empty:
        return pd.DataFrame()

    table = pd.crosstab(
        temp["Strategie"],
        temp[second_name]
    )

    return table


strategie_topic = make_cooccurrence(
    df,
    "_topics",
    "Topic"
)

strategie_actor = make_cooccurrence(
    df,
    "_actors",
    "Actor"
)

strategie_visual = make_cooccurrence(
    df,
    "_visuals",
    "Visual"
)

strategie_audio = make_cooccurrence(
    df,
    "_audios",
    "Audio"
)


strategie_topic.to_csv(
    DATEN_ORDNER / "09_strategie_topic.csv",
    encoding="utf-8-sig"
)

strategie_actor.to_csv(
    DATEN_ORDNER / "10_strategie_actor.csv",
    encoding="utf-8-sig"
)

strategie_visual.to_csv(
    DATEN_ORDNER / "11_strategie_visual.csv",
    encoding="utf-8-sig"
)

strategie_audio.to_csv(
    DATEN_ORDNER / "12_strategie_audio.csv",
    encoding="utf-8-sig"
)


# ============================================================
# 20. STRATEGIEKOMBINATIONEN
# ============================================================

combination_counter = Counter()

for strategies in df["_strategies"]:

    if strategies:

        combination = " + ".join(
            sorted(strategies)
        )

        combination_counter[combination] += 1


strategie_kombinationen = pd.DataFrame(
    combination_counter.items(),
    columns=[
        "Strategiekombination",
        "Anzahl_Segmente"
    ]
)

if not strategie_kombinationen.empty:

    strategie_kombinationen = (
        strategie_kombinationen.sort_values(
            "Anzahl_Segmente",
            ascending=False
        )
    )


strategie_kombinationen.to_csv(
    DATEN_ORDNER / "13_strategiekombinationen.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 21. STRATEGIEDICHTE
# ============================================================

df_density = df[
    [
        "Folge",
        "Jahr",
        "Segment_Nr",
        "Segment_ID",
        "Dauer_Sekunden"
    ]
].copy()

df_density["Anzahl_Strategien"] = (
    df["_strategies"].apply(len)
)

strategiedichte_jahr = (
    df_density.groupby("Jahr")
    .agg(
        Segmente=("Segment_ID", "count"),
        Mittelwert_Strategien=("Anzahl_Strategien", "mean"),
        Median_Strategien=("Anzahl_Strategien", "median"),
        Maximum_Strategien=("Anzahl_Strategien", "max"),
    )
    .reset_index()
)

strategiedichte_jahr.to_csv(
    DATEN_ORDNER / "14_strategiedichte_pro_jahr.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 22. SEGMENTE MIT / OHNE STRATEGIE
# ============================================================

df_presence = df[
    [
        "Folge",
        "Jahr",
        "Segment_ID"
    ]
].copy()

df_presence["Strategie_vorhanden"] = (
    df["_strategies"].apply(
        lambda x: "Mit Strategie"
        if len(x) > 0
        else "Ohne Strategie"
    )
)

strategie_presence = (
    df_presence.groupby(
        ["Jahr", "Strategie_vorhanden"]
    )
    .size()
    .unstack(fill_value=0)
    .reset_index()
)

strategie_presence.to_csv(
    DATEN_ORDNER / "15_segmente_mit_ohne_strategie.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 23. STATISTIK PRO FOLGE
# ============================================================

folgenstatistik = (
    df.groupby(
        ["Jahr", "Folge"]
    )
    .agg(
        Segmente=("Segment_ID", "count"),
        Gesamtdauer_Sekunden=("Dauer_Sekunden", "sum"),
        Durchschnitt_Segment_Sekunden=("Dauer_Sekunden", "mean"),
        Median_Segment_Sekunden=("Dauer_Sekunden", "median"),
    )
    .reset_index()
)

folgenstatistik["Gesamtdauer_Minuten"] = (
    folgenstatistik["Gesamtdauer_Sekunden"] / 60
)

folgenstatistik.to_csv(
    DATEN_ORDNER / "16_folgenstatistik.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 24. EINZELNE SEGMENTLÄNGEN MIT STRATEGIE
# ============================================================

if not strategy_long.empty:

    einzelne_segmentlaengen = strategy_long[
        [
            "Folge",
            "Jahr",
            "Segment_Nr",
            "Segment_ID",
            "Strategie",
            "Dauer_Sekunden",
            "Dauer_Minuten",
        ]
    ].copy()

else:
    einzelne_segmentlaengen = pd.DataFrame()


einzelne_segmentlaengen.to_csv(
    DATEN_ORDNER / "17_einzelne_segmentlaengen.csv",
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# 25. EXCEL-GESAMTAUSGABE
# ============================================================

excel_path = (
    ERGEBNISSE_ORDNER /
    "Masterarbeit_Auswertung.xlsx"
)

with pd.ExcelWriter(
    excel_path,
    engine="openpyxl"
) as writer:

    df_public.to_excel(
        writer,
        sheet_name="Segmente",
        index=False
    )

    korpusstatistik.to_excel(
        writer,
        sheet_name="Korpusstatistik",
        index=False
    )

    strategien_gesamt.to_excel(
        writer,
        sheet_name="Strategien_Gesamt",
        index=False
    )

    strategien_jahr.to_excel(
        writer,
        sheet_name="Strategien_Jahr",
        index=False
    )

    strategie_relativ.to_excel(
        writer,
        sheet_name="Strategien_Relativ",
        index=False
    )

    strategie_runtime.to_excel(
        writer,
        sheet_name="Strategien_Laufzeit",
        index=False
    )

    strategie_segmentlaenge.to_excel(
        writer,
        sheet_name="Laenge_Strategie",
        index=False
    )

    strategie_segmentlaenge_jahr.to_excel(
        writer,
        sheet_name="Laenge_Strategie_Jahr",
        index=False
    )

    strategie_topic.to_excel(
        writer,
        sheet_name="Strategie_Topic"
    )

    strategie_actor.to_excel(
        writer,
        sheet_name="Strategie_Actor"
    )

    strategie_visual.to_excel(
        writer,
        sheet_name="Strategie_Visual"
    )

    strategie_audio.to_excel(
        writer,
        sheet_name="Strategie_Audio"
    )

    strategie_kombinationen.to_excel(
        writer,
        sheet_name="Kombinationen",
        index=False
    )

    strategiedichte_jahr.to_excel(
        writer,
        sheet_name="Strategiedichte",
        index=False
    )

    strategie_presence.to_excel(
        writer,
        sheet_name="Mit_ohne_Strategie",
        index=False
    )

    folgenstatistik.to_excel(
        writer,
        sheet_name="Folgenstatistik",
        index=False
    )

    einzelne_segmentlaengen.to_excel(
        writer,
        sheet_name="Segmentlaengen_Einzeln",
        index=False
    )


# ============================================================
# 26. GRAFIK 1:
# HÄUFIGKEIT DER PROPAGANDASTRATEGIEN
# ============================================================

if not strategien_gesamt.empty:

    plot_data = (
        strategien_gesamt
        .sort_values(
            "Anzahl_Segmente",
            ascending=True
        )
    )

    fig, ax = plt.subplots(
        figsize=(10, 6)
    )

    ax.barh(
        plot_data["Strategie"],
        plot_data["Anzahl_Segmente"]
    )

    ax.set_xlabel("Anzahl der Segmente")
    ax.set_ylabel("Propagandastrategie")

    ax.set_title(
        "Häufigkeit der Propagandastrategien im Gesamtkorpus"
    )

    ax.grid(
        axis="x",
        alpha=0.25
    )

    plt.tight_layout()

    plt.savefig(
        ABBILDUNGEN_ORDNER /
        "01_strategien_gesamt.png",
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()


# ============================================================
# 27. GRAFIK 2:
# RELATIVE STRATEGIEHÄUFIGKEIT 1940–1945
# ============================================================

if not strategie_relativ.empty:

    pivot_relativ = (
        strategie_relativ.pivot(
            index="Jahr",
            columns="Strategie",
            values="Anteil_Segmente_Prozent"
        )
        .fillna(0)
        .sort_index()
    )

    fig, ax = plt.subplots(
        figsize=(12, 7)
    )

    for strategy in pivot_relativ.columns:

        ax.plot(
            pivot_relativ.index,
            pivot_relativ[strategy],
            marker="o",
            label=strategy
        )

    ax.set_xlabel("Jahr")
    ax.set_ylabel(
        "Anteil der Segmente (%)"
    )

    ax.set_title(
        "Entwicklung der Propagandastrategien 1940–1945"
    )

    ax.legend(
        bbox_to_anchor=(1.02, 1),
        loc="upper left"
    )

    ax.grid(alpha=0.25)

    plt.tight_layout()

    plt.savefig(
        ABBILDUNGEN_ORDNER /
        "02_strategien_zeitverlauf.png",
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()


# ============================================================
# 28. GRAFIK 3:
# LAUFZEITANTEIL DER STRATEGIEN PRO JAHR
# ============================================================

if not strategie_runtime.empty:

    pivot_runtime = (
        strategie_runtime.pivot(
            index="Jahr",
            columns="Strategie",
            values="Anteil_Laufzeit_Prozent"
        )
        .fillna(0)
        .sort_index()
    )

    fig, ax = plt.subplots(
        figsize=(12, 7)
    )

    for strategy in pivot_runtime.columns:

        ax.plot(
            pivot_runtime.index,
            pivot_runtime[strategy],
            marker="o",
            label=strategy
        )

    ax.set_xlabel("Jahr")
    ax.set_ylabel(
        "Anteil der annotierten Laufzeit (%)"
    )

    ax.set_title(
        "Laufzeitanteile der Propagandastrategien 1940–1945"
    )

    ax.legend(
        bbox_to_anchor=(1.02, 1),
        loc="upper left"
    )

    ax.grid(alpha=0.25)

    plt.tight_layout()

    plt.savefig(
        ABBILDUNGEN_ORDNER /
        "03_strategien_laufzeit.png",
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()


# ============================================================
# 29. GRAFIK 4:
# BOXPLOT SEGMENTLÄNGE PRO STRATEGIE
# ============================================================

if not strategy_long.empty:

    strategy_order = (
        strategy_long.groupby("Strategie")[
            "Dauer_Sekunden"
        ]
        .median()
        .sort_values()
        .index
        .tolist()
    )

    box_data = [
        strategy_long.loc[
            strategy_long["Strategie"] == strategy,
            "Dauer_Sekunden"
        ].dropna().values
        for strategy in strategy_order
    ]

    fig, ax = plt.subplots(
        figsize=(11, 7)
    )

    ax.boxplot(
    box_data,
    tick_labels=strategy_order,
    vert=False,
    showfliers=True
)
    ax.set_xlabel(
        "Segmentlänge (Sekunden)"
    )

    ax.set_ylabel(
        "Propagandastrategie"
    )

    ax.set_title(
        "Verteilung der Segmentlängen nach Propagandastrategie"
    )

    ax.grid(
        axis="x",
        alpha=0.25
    )

    plt.tight_layout()

    plt.savefig(
        ABBILDUNGEN_ORDNER /
        "04_segmentlaengen_boxplot.png",
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()


# ============================================================
# 30. GRAFIK 5:
# MEDIANE SEGMENTLÄNGE NACH JAHR
# ============================================================

if not strategie_segmentlaenge_jahr.empty:

    pivot_median = (
        strategie_segmentlaenge_jahr.pivot(
            index="Jahr",
            columns="Strategie",
            values="Median_Sekunden"
        )
        .sort_index()
    )

    fig, ax = plt.subplots(
        figsize=(12, 7)
    )

    for strategy in pivot_median.columns:

        ax.plot(
            pivot_median.index,
            pivot_median[strategy],
            marker="o",
            label=strategy
        )

    ax.set_xlabel("Jahr")
    ax.set_ylabel(
        "Mediane Segmentlänge (Sekunden)"
    )

    ax.set_title(
        "Mediane Segmentlänge nach Propagandastrategie"
    )

    ax.legend(
        bbox_to_anchor=(1.02, 1),
        loc="upper left"
    )

    ax.grid(alpha=0.25)

    plt.tight_layout()

    plt.savefig(
        ABBILDUNGEN_ORDNER /
        "05_segmentlaenge_trend.png",
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()


# ============================================================
# 31. HEATMAP-FUNKTION
# ============================================================
#
# Ohne seaborn, damit nur matplotlib benötigt wird.
# ============================================================

def save_heatmap(
    table,
    title,
    xlabel,
    ylabel,
    filename
):

    if table.empty:
        return

    # Größe dynamisch an Anzahl der Kategorien anpassen
    width = max(
        8,
        len(table.columns) * 0.8
    )

    height = max(
        5,
        len(table.index) * 0.65
    )

    fig, ax = plt.subplots(
        figsize=(width, height)
    )

    matrix = table.values

    image = ax.imshow(
        matrix,
        aspect="auto",
        cmap="Blues"
    )

    ax.set_xticks(
        range(len(table.columns))
    )

    ax.set_xticklabels(
        table.columns,
        rotation=45,
        ha="right"
    )

    ax.set_yticks(
        range(len(table.index))
    )

    ax.set_yticklabels(
        table.index
    )

    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.set_title(title)

    # Werte in die Zellen schreiben
    max_value = (
        matrix.max()
        if matrix.size > 0
        else 0
    )

    for i in range(matrix.shape[0]):

        for j in range(matrix.shape[1]):

            value = matrix[i, j]

            if value > 0:

                ax.text(
                    j,
                    i,
                    str(int(value)),
                    ha="center",
                    va="center",
                    fontsize=8
                )

    fig.colorbar(
        image,
        ax=ax,
        label="Anzahl der Segmente"
    )

    plt.tight_layout()

    plt.savefig(
        ABBILDUNGEN_ORDNER / filename,
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()


# ============================================================
# 32. GRAFIK 6:
# STRATEGIE × TOPIC
# ============================================================

save_heatmap(
    strategie_topic,
    "Propagandastrategien und Themen",
    "Topic",
    "Propagandastrategie",
    "06_strategie_topic.png"
)


# ============================================================
# 33. GRAFIK 7:
# STRATEGIE × AUDIO
# ============================================================

save_heatmap(
    strategie_audio,
    "Propagandastrategien und auditive Gestaltung",
    "Audio",
    "Propagandastrategie",
    "07_strategie_audio.png"
)


# ============================================================
# 34. GRAFIK 8:
# STRATEGIE × VISUAL
# ============================================================

save_heatmap(
    strategie_visual,
    "Propagandastrategien und visuelle Gestaltung",
    "Visuelle Technik",
    "Propagandastrategie",
    "08_strategie_visual.png"
)


# ============================================================
# 35. GRAFIK 9:
# STRATEGIE × ACTOR
# ============================================================

save_heatmap(
    strategie_actor,
    "Propagandastrategien und Akteure",
    "Actor",
    "Propagandastrategie",
    "09_strategie_actor.png"
)


# ============================================================
# 36. GRAFIK 10:
# HÄUFIGSTE STRATEGIEKOMBINATIONEN
# ============================================================

if not strategie_kombinationen.empty:

    top_combinations = (
        strategie_kombinationen
        .head(15)
        .sort_values(
            "Anzahl_Segmente",
            ascending=True
        )
    )

    fig, ax = plt.subplots(
        figsize=(11, 8)
    )

    ax.barh(
        top_combinations["Strategiekombination"],
        top_combinations["Anzahl_Segmente"]
    )

    ax.set_xlabel(
        "Anzahl der Segmente"
    )

    ax.set_ylabel(
        "Strategiekombination"
    )

    ax.set_title(
        "Häufigste Kombinationen von Propagandastrategien"
    )

    ax.grid(
        axis="x",
        alpha=0.25
    )

    plt.tight_layout()

    plt.savefig(
        ABBILDUNGEN_ORDNER /
        "10_strategiekombinationen.png",
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()


# ============================================================
# 37. GRAFIK 11:
# SEGMENTE MIT / OHNE STRATEGIE
# ============================================================

if not strategie_presence.empty:

    plot_presence = (
        strategie_presence
        .set_index("Jahr")
    )

    fig, ax = plt.subplots(
        figsize=(9, 6)
    )

    plot_presence.plot(
        kind="bar",
        stacked=True,
        ax=ax
    )

    ax.set_xlabel("Jahr")
    ax.set_ylabel(
        "Anzahl der Segmente"
    )

    ax.set_title(
        "Segmente mit und ohne annotierte Propagandastrategie"
    )

    ax.legend(
        title=""
    )

    ax.grid(
        axis="y",
        alpha=0.25
    )

    plt.xticks(
        rotation=0
    )

    plt.tight_layout()

    plt.savefig(
        ABBILDUNGEN_ORDNER /
        "11_mit_ohne_strategie.png",
        dpi=300,
        bbox_inches="tight"
    )

    plt.close()


# ============================================================
# 38. GRAFIK 12:
# ANZAHL SEGMENTE PRO JAHR
# ============================================================

segments_year_plot = (
    df.groupby("Jahr")
    .size()
    .sort_index()
)

fig, ax = plt.subplots(
    figsize=(8, 5)
)

ax.bar(
    segments_year_plot.index.astype(str),
    segments_year_plot.values
)

ax.set_xlabel("Jahr")
ax.set_ylabel("Anzahl der Segmente")

ax.set_title(
    "Umfang des Annotationskorpus nach Jahr"
)

ax.grid(
    axis="y",
    alpha=0.25
)

plt.tight_layout()

plt.savefig(
    ABBILDUNGEN_ORDNER /
    "12_segmente_pro_jahr.png",
    dpi=300,
    bbox_inches="tight"
)

plt.close()


# ============================================================
# 39. QUALITÄTSKONTROLLE
# ============================================================

print()
print("=" * 70)
print("QUALITÄTSKONTROLLE")
print("=" * 70)

print(
    f"EAF-Dateien: {len(eaf_files)}"
)

print(
    f"Folgen erkannt: {df['Folge'].nunique()}"
)

print(
    f"Segmente insgesamt: {len(df)}"
)

print(
    f"Gesamtdauer der Segmente: "
    f"{df['Dauer_Minuten'].sum():.2f} Minuten"
)

print()

print("Folgen nach Jahr:")

print(
    df.groupby("Jahr")["Folge"]
    .nunique()
)

print()

print("Segmente nach Jahr:")

print(
    df.groupby("Jahr")
    .size()
)

print()

print("Propagandastrategien:")

if not strategien_gesamt.empty:

    print(
        strategien_gesamt[
            [
                "Strategie",
                "Anzahl_Segmente"
            ]
        ].to_string(index=False)
    )

else:
    print(
        "Keine Propagandastrategien gefunden."
    )


# ============================================================
# 40. KATEGORIEN AUSGEBEN
# ============================================================
#
# Sehr wichtig für die Kontrolle auf Tippfehler.
# Wenn z.B. "Heroisierung" und "Heroisierunng" erscheinen,
# kannst du die EAF korrigieren oder oben in der
# NORMALISIERUNG ergänzen.
# ============================================================

def print_unique(title, values):

    print()
    print(title)
    print("-" * len(title))

    for value in sorted(values):
        print(value)


print_unique(
    "Gefundene Propagandastrategien",
    set(
        value
        for values in df["_strategies"]
        for value in values
    )
)

print_unique(
    "Gefundene Topics",
    set(
        value
        for values in df["_topics"]
        for value in values
    )
)

print_unique(
    "Gefundene Actors",
    set(
        value
        for values in df["_actors"]
        for value in values
    )
)

print_unique(
    "Gefundene Visual Techniques",
    set(
        value
        for values in df["_visuals"]
        for value in values
    )
)

print_unique(
    "Gefundene Audio-Kategorien",
    set(
        value
        for values in df["_audios"]
        for value in values
    )
)


# ============================================================
# 41. ABSCHLUSS
# ============================================================

print()
print("=" * 70)
print("AUSWERTUNG ABGESCHLOSSEN")
print("=" * 70)

print(
    f"Ausgabeordner:\n{AUSGABE_ORDNER}"
)

print()

print(
    f"Excel-Datei:\n{excel_path}"
)

print()

print(
    "CSV-Dateien befinden sich in:"
)

print(
    DATEN_ORDNER
)

print()

print(
    "Abbildungen befinden sich in:"
)

print(
    ABBILDUNGEN_ORDNER
)

print("=" * 70)