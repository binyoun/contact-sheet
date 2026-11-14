#!/usr/bin/env python3
"""Build books.js from the RMIT Vietnam library export of photography books.

Usage (from the repo root):
  python3 scripts/build_books.py ~/photography_books_library_rmit_oct_2025.xlsx

Refresh each semester: export the same Alma report again and rerun this script.
Only public catalogue facts are kept (title, author, year, call number, catalogue ID);
barcodes, item IDs and loan status are dropped because they go stale or are not useful.

To link a book to a photographer on the site, add a line to ARTIST_BOOKS below.
"""
import json, re, sys
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent.parent

# (pattern searched in "title + creator", artist name exactly as in artists.js)
ARTIST_BOOKS = [
    (r"Hanabi|Fundamental cycles", "Rinko Kawauchi"),
    (r"^The nature of photographs|^Stephen Shore /", "Stephen Shore"),
    (r"Dawoud Bey", "Dawoud Bey"),
    (r"McCurry", "Steve McCurry"),
    (r"^Eyes open : 23", "Susan Meiselas"),
    (r"Alec Soth|Gus the Great", "Alec Soth"),
    (r"Richard Misrach", "Richard Misrach"),
    (r"Thomas Demand", "Thomas Demand"),
    (r"Daido Moriyama", "Daido Moriyama"),
    (r"Burtynsky", "Edward Burtynsky"),
    (r"Candida H", "Candida Höfer"),
    (r"Crewdson", "Gregory Crewdson"),
    (r"ballad of sexual dependency", "Nan Goldin"),
    (r"Erwin Olaf", "Erwin Olaf"),
    (r"Leibovitz", "Annie Leibovitz"),
    (r"Thomas Ruff", "Thomas Ruff"),
    (r"Tracey Moffatt", "Tracey Moffatt"),
    (r"Bristol, England, June|The photobook : a history", "Martin Parr"),
    (r"Gursky", "Andreas Gursky"),
    (r"Bialobrzeski", "Peter Bialobrzeski"),
    (r"^Shirin Neshat", "Shirin Neshat"),
    (r"Rineke Dijkstra", "Rineke Dijkstra"),
    (r"Hank Willis Thomas", "Hank Willis Thomas"),
    (r"D'Agata", "Antoine d'Agata"),
    (r"Juergen Teller", "Juergen Teller"),
    (r"Dinh Q", "Dinh Q. Lê"),
    (r"Doisneau", "Robert Doisneau"),
    (r"Helene Binet", "Helene Binet"),
    (r"Irving Penn", "Irving Penn"),
    (r"Avedon", "Richard Avedon"),
    (r"Tim Walker", "Tim Walker"),
    (r"^Babe / edited by Petra Collins", "Petra Collins"),
    (r"Herb Ritts", "Herb Ritts"),
    (r"^Saul Leiter", "Saul Leiter"),
    (r"^Lorna Simpson", "Lorna Simpson"),
    (r"Julius Shulman", "Julius Shulman"),
    (r"Ansel Adams", "Ansel Adams"),
    (r"Morimura", "Yasumasa Morimura"),
    (r"Sally Mann", "Sally Mann"),
    (r"Wolfgang Tillmans", "Wolfgang Tillmans"),
    (r"^Cindy Sherman", "Cindy Sherman"),
    (r"Chaotic harmony", "Koo, Bohnchang"),
    (r"Documentary &\W*anti-graphic", "Henri Cartier-Bresson"),
    (r"Documentary &\W*anti-graphic|Walker Evans", "Walker Evans"),
    (r"Diane Arbus", "Diane Arbus"),
    (r"Robert Capa", "Robert Capa"),
    (r"Brassa", "Brassaï"),
    (r"Lee Miller", "Lee Miller"),
    (r"Atget", "Eugène Atget"),
    (r"^Helen Levitt", "Helen Levitt"),
    (r"Mary Ellen Mark", "Mary Ellen Mark"),
    (r"^Our farm", "Raymond Depardon"),
    (r"Fukase", "Masahisa Fukase"),
    (r"Helmut Newton", "Helmut Newton"),
    (r"Lindbergh", "Peter Lindbergh"),
    (r"LaChapelle", "David LaChapelle"),
    (r"Rankin", "Rankin"),
    (r"Mert Alas", "Mert Alas and Marcus Piggott"),
    (r"sartorialist", "Scott Schuman"),
    (r"Nadav Kander", "Nadav Kander"),
    (r"Martin Schoeller", "Martin Schoeller"),
    (r"transparent city", "Michael Wolf"),
    (r"Kehinde Wiley", "Kehinde Wiley"),
    (r"Garry Fabian Miller", "Garry Fabian Miller"),
    (r"^Along some rivers", "Robert Adams"),
    (r"Bourke-White", "Margaret Bourke-White"),
    (r"Steichen", "Edward Steichen"),
    (r"Feininger", "Andreas Feininger"),
    (r"Gedney", "William Gedney"),
    (r"Goldsworthy", "Andy Goldsworthy"),
    (r"Roni Horn", "Roni Horn"),
    (r"Morrisroe", "Mark Morrisroe"),
    (r"Toledano", "Phillip Toledano"),
    (r"Louis Faurer", "Louis Faurer"),
    (r"Peter Beard", "Peter Beard"),
    (r"Giacomelli", "Mario Giacomelli"),
    (r"Ralph Gibson", "Ralph Gibson"),
    (r"ZZYZX", "Gregory Halpern"),
    (r"^Alison Jackson", "Alison Jackson"),
    (r"Heinrich Ku", "Heinrich Kühn"),
    (r"Bertien van Manen", "Bertien van Manen"),
    (r"Chris Marker", "Chris Marker"),
    (r"Rozovsky", "Irina Rozovsky"),
    (r"Tich[ýy]", "Miroslav Tichý"),
    (r"Christopher Williams|^Program : for example", "Christopher Williams"),
    (r"Martha Rosler", "Martha Rosler"),
    (r"Stieglitz", "Alfred Stieglitz"),
    (r"Liu Zheng", "Liu Zheng"),
    (r"Izima", "Izima Kaoru"),
    (r"^Fruits /", "Shoichi Aoki"),
    (r"Tenneson", "Joyce Tenneson"),
    (r"Wegman", "William Wegman"),
    (r"Darren Almond", "Darren Almond"),
    (r"Horst P\. Horst", "Horst P. Horst"),
    (r"Blumenfeld", "Erwin Blumenfeld"),
    (r"Shambroom", "Paul Shambroom"),
    (r"Basilico", "Gabriele Basilico"),
    (r"Dujardin", "Filip Dujardin"),
    (r"Hamann", "Horst Hamann"),
    (r"Frederick H\. Evans", "Frederick H. Evans"),
    (r"Edward Weston", "Edward Weston"),
    (r"David Bailey", "David Bailey"),
    (r"Blinked Myself Awake", "Bieke Depooter"),
    (r"In Sook Kim", "In Sook Kim"),
    (r"^Yang Yong", "Yang Yong"),
    (r"Sigismondi", "Floria Sigismondi"),
    (r"Mario Sorrenti", "Mario Sorrenti"),
    (r"atlas of beauty", "Mihaela Noroc"),
    (r"Touch me I.m sick", "Charles Peterson"),
    (r"Paiva", "Troy Paiva"),
    (r"Meyerowitz", "Joel Meyerowitz"),
    (r"Tahara", "Keiichi Tahara"),
    (r"Tropical gift", "Christian Lutz"),
    (r"Kosmos", "Adam Bartos"),
    (r"Oltheten", "Paulien Oltheten"),
    (r"Kawashima", "Kotori Kawashima"),
    (r"Greg Gorman", "Greg Gorman"),
    (r"Barbieri", "Gian Paolo Barbieri"),
    (r"Dior couture", "Patrick Demarchelier"),
    (r"Brodovitch", "Alexey Brodovitch"),
    (r"Selby", "Todd Selby"),
    (r"^Advanced style", "Ari Seth Cohen"),
    (r"transportation of place", "Andrea Robbins and Max Becher"),
    (r"Louise Lawler", "Louise Lawler"),
    (r"Klutsis", "Gustav Klutsis"),
    (r"Yoon Ji Seon", "Yoon, Ji Seon"),
]

HOWTO = re.compile(r"\b(how to|guide|techniques?|tips|masterclass|posing|for dummies|made simple|for fun|handbook|workshop|"
                   r"classroom|field guide|lighting|mastering|problem solver|course|step[- ]by[- ]step|secrets|portfolio|"
                   r"business|legal forms|lightroom|photoshop|illustrator|exposure|composition|diy|projects|photographing|"
                   r"photographer's|for photographers|essential|complete|manual|learn|basic|digital photography|the art and style)\b", re.I)

TAGS = [  # (Contact Sheet tag key, pattern)
    ("portrait", r"portrait|people|faces?\b|posing|pose|headshot|twins"),
    ("landscape", r"landscape|nature|panorama|wilderness|garden|earth|american west|rivers"),
    ("street", r"street|urban|city"),
    ("stilllife", r"still life|objects?\b"),
    ("fashion", r"fashion|vogue|style\b|couture|sartorial|beauty|glamour|prada|mcqueen"),
    ("food", r"food|hungry|cook|culinary|recipe"),
    ("product", r"product|advertising|commercial|marketing"),
    ("architecture", r"architect|interior|buildings?|vertical|modernism|libraries|homes?\b"),
    ("documentary", r"documentary|social history|reportage|magnum|stories"),
    ("photojournalism", r"photojournalis|news|war\b|conflict|violence|magnum"),
    ("lifestyle", r"lifestyle|wedding|travel"),
    ("moving", r"cinema|film-?making|video|movie|moving image"),
    ("abstract", r"abstract"),
    ("diary", r"snapshot|polaroid|diary"),
    ("studio", r"studio|lighting|flash|\blight\b|strobe"),
    ("bw", r"black (and|&) white|monochrome|b&w|zone system"),
    ("night", r"night|moon"),
    ("altprocess", r"darkroom|alternative process|cyanotype|pinhole|polaroid|film photography|analog"),
    ("composite", r"photoshop|composit|retouch|manipulat"),
    ("color", r"colou?r"),
    ("aerial", r"drone|aerial"),
    ("lightpaint", r"light painting|painting with light"),
    ("collage", r"collage|montage"),
]

CALL_TAGS = [("746.9", "fashion"), ("391", "fashion"), ("779.97", "fashion"), ("779.2", "portrait"), ("778.92", "portrait"),
             ("72", "architecture"), ("779.4", "landscape"), ("779.3", "landscape"), ("778.93", "landscape"),
             ("778.99", "architecture"), ("779.96", "food"), ("778.72", "studio"), ("771.4", "studio")]


def clean_title(t):
    t = (t or "").replace("​", "").strip()
    t = re.split(r"\s+/\s*", t)[0]  # drop the statement of responsibility
    t = re.sub(r"\[|\]", "", t)
    t = re.sub(r"\s+:\s*", ": ", t).strip(" .:;,")
    return t[:1].upper() + t[1:]


def creator(imprint):
    m = re.search(r"By (.+?)\s*\(", imprint or "")
    if not m:
        return ""
    c = re.sub(r",?\s*\d{4}-(\d{4})?", "", m.group(1)).strip(" ,.")
    if "," in c:
        last, first = [p.strip(" ,.") for p in c.split(",", 1)]
        if first and not re.search(r"\d", first):
            c = f"{first} {last}"
    return c


def year(imprint):
    ys = re.findall(r"(1[89]\d\d|20\d\d)", (imprint or "").split("(")[-1])
    return int(ys[-1]) if ys else None


def kind(callno, text, linked):
    if linked:
        return "artist"
    d = callno[:3]
    if d in {"771", "772", "773", "774", "775", "776", "778", "006", "621", "659", "522"} or callno.startswith(("741.6", "745.5", "746.6", "708.4")):
        return "how"
    if HOWTO.search(text) and not callno.startswith(("779.0", "779.09")):
        return "how"
    if re.search(r"marketing|how photographs are sold|instagram|making a life|living in photography", text, re.I):
        return "how"
    if re.search(r"\bhistory\b|encyclopedia|dictionary|most influential|key concepts|essays|theory|reader\b", text, re.I):
        return "ideas"
    if re.match(r"770(\.[01]|\.9(?!2)|$| )", callno) or d in {"070", "303", "001", "363", "791", "759", "701", "808"}:
        return "ideas"
    return "collection"


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    ws = openpyxl.load_workbook(Path(sys.argv[1]).expanduser(), read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h or "") for h in rows[0]]
    books, copies = {}, {}
    known = set(re.findall(r'^A\("([^"]+)"', (ROOT / "artists.js").read_text(encoding="utf-8"), re.M))
    unknown = {n for _, n in ARTIST_BOOKS} - known
    if unknown:
        sys.exit(f"ARTIST_BOOKS names not in artists.js: {sorted(unknown)}")
    for r in rows[1:]:
        d = dict(zip(head, r))
        mms = str(d.get("MMS ID") or "").strip()
        if not mms:
            continue
        copies[mms] = copies.get(mms, 0) + 1
        if mms in books:
            continue
        raw_title, imprint = d.get("Title") or "", d.get("Type / Creator / Imprint") or ""
        callno = re.sub(r"\s+", " ", str(d.get("Call Number") or "").replace("​", "")).strip()
        text = f"{raw_title} {imprint}"
        artists = sorted({name for pat, name in ARTIST_BOOKS if re.search(pat, text, re.I)})
        title = clean_title(raw_title)
        k = kind(callno, f"{title} {raw_title}", bool(artists))
        tags = {key for key, pat in TAGS if re.search(pat, raw_title, re.I)}
        for prefix, key in CALL_TAGS:  # the shelf position says what the book is about
            if callno.startswith(prefix):
                tags.add(key)
        if callno.startswith("704.942 1"):  # figure-drawing reference books: keep on the shelf, never suggest
            tags = set()
        tags = sorted(tags)
        b = {"id": mms, "t": title, "by": creator(imprint), "y": year(imprint), "cn": callno, "k": k}
        if tags:
            b["g"] = tags
        if artists:
            b["a"] = artists
        books[mms] = b
    out = sorted(books.values(), key=lambda b: (b["cn"], b["t"]))
    for b in out:
        if copies[b["id"]] > 1:
            b["n"] = copies[b["id"]]
    js = ("// Generated by scripts/build_books.py from the RMIT Vietnam library export. Do not edit by hand.\n"
          f"// {len(out)} titles, Beanland Library (Saigon South).\n"
          "const LIBRARY = { name: 'Beanland Library, Saigon South', updated: 'October 2025', "
          "catalogue: (id) => 'https://rmitlibraryvn.rmit.edu.vn/discovery/fulldisplay?context=L&vid=84RVI_INST:84RVI&docid=alma' + id };\n"
          "const BOOKS = " + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";\n")
    (ROOT / "books.js").write_text(js, encoding="utf-8")
    count = {}
    for b in out:
        count[b["k"]] = count.get(b["k"], 0) + 1
    linked = {n for b in out for n in b.get("a", [])}
    print(f"wrote books.js: {len(out)} titles {count}; {len(linked)} photographers linked")
    missing = sorted({n for _, n in ARTIST_BOOKS} - linked)
    if missing:
        print("patterns that matched no book:", missing)


if __name__ == "__main__":
    main()
