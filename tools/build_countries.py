"""Build references/countries.csv from the IPBES IBOK geography file.

Usage (Programme Assistant):
  1. Download https://ibok.ipbes.net/rest/v1/geo/data?format=application/rdf%2Bxml
     and save it as ipbes-geo.rdf
  2. python build_countries.py ipbes-geo.rdf ../skills/ipbes-editorial/references/countries.csv
Standard library only.
"""
import csv, sys, unicodedata, xml.etree.ElementTree as ET

NS = {
    "rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
    "skos": "http://www.w3.org/2004/02/skos/core#",
    "g": "https://ibok.ipbes.net/geography/",
}
RDF_ABOUT = "{%s}about" % NS["rdf"]
RDF_RES = "{%s}resource" % NS["rdf"]
XML_LANG = "{http://www.w3.org/XML/1998/namespace}lang"
REGIONS = {"African States", "Asia-Pacific States", "Eastern European States",
           "Latin American and Caribbean States", "Western European and other States"}


def clean(s):
    """Trim stray whitespace/newlines and use curly apostrophes (house style)."""
    return " ".join(s.split()).replace("'", "’")


def en(el, tag):
    vals = [clean(x.text) for x in el.findall(tag, NS) if x.get(XML_LANG) == "en" and x.text]
    return vals


def main(src, dst):
    root = ET.parse(src).getroot()
    descs = root.findall("rdf:Description", NS)
    collections, concepts = {}, {}
    for d in descs:
        about = d.get(RDF_ABOUT)
        types = {t.get(RDF_RES) for t in d.findall("rdf:type", NS)}
        if d.find("skos:notation", NS) is not None:
            concepts[about] = d
        members = [m.get(RDF_RES) for m in d.findall("skos:member", NS)]
        label = (en(d, "skos:prefLabel") or [None])[0]
        if members or any(t and t.endswith("#Collection") for t in types):
            c = collections.setdefault(about, {"label": None, "members": []})
            c["label"] = c["label"] or label
            c["members"] += members
    region_of, subregion_of = {}, {}
    for c in collections.values():
        target = region_of if c["label"] in REGIONS else subregion_of
        for m in c["members"]:
            if c["label"]:
                target.setdefault(m, c["label"])
    rows = []
    for about, d in concepts.items():
        pref = (en(d, "skos:prefLabel") or [""])[0]
        short = pref[:-6] if pref.endswith(" (the)") else pref
        alt = "; ".join(en(d, "skos:altLabel"))
        iso3 = d.find("skos:notation", NS).text
        ms = d.find("g:memberStatus", NS)
        since = d.find("g:memberSince", NS)
        rows.append({
            "name_in_text": short,
            "unterm_short_name": pref,
            "formal_name": alt,
            "iso3": iso3,
            "ipbes_member": (ms.text if ms is not None else ""),
            "member_since": (since.text[:10] if since is not None and since.text else ""),
            "un_regional_group": region_of.get(about, ""),
            "ipbes_subregion": subregion_of.get(about, ""),
        })
    # English alphabetical order, ignoring accents (Côte d’Ivoire before Croatia; Türkiye before Turkmenistan)
    rows.sort(key=lambda r: unicodedata.normalize("NFKD", r["name_in_text"]).encode("ascii", "ignore").decode().lower())
    with open(dst, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    print(f"{len(rows)} countries written; {sum(r['ipbes_member']=='true' for r in rows)} IPBES members")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
