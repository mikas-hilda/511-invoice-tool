# 5.11 Invoice PDF -> Excel įrankis

Šitas įrankis konvertuoja 5.11 PDF sąskaitas į `.xlsx` pagal tavo nustatytą struktūrą.

## Ką daro

- Ištraukia:
  - Invoice Number
  - Invoice Date
  - Customer PO
  - INVOICE TOTAL
- Ištraukia prekių lentelės eilutes pagal faktines PDF lentelės stulpelių ribas.
- Sugeneruoja Excel su:
  - antrašte 1-2 eilutėse;
  - tuščia 3 eilute;
  - prekių lentele nuo 4 eilutės.
- Validuoja:
  - PDF eilučių skaičius = Excel eilučių skaičius;
  - Amount suma = INVOICE TOTAL;
  - Qty, Your Unit Price, Amount yra skaitiniai;
  - Color lieka Color stulpelyje;
  - M stulpelyje pašalinamas tik galinis `ABR`;
  - D stulpelis = `5-M-spalvos_kodas`.

## Diegimas

### 1. Įsidiek Python

Atsisiųsk iš:
https://www.python.org/downloads/

Diegiant pažymėk:
`Add Python to PATH`

### 2. Išarchyvuok šį ZIP

Pvz. į folderį:

`C:\511_invoice_tool`

arba Mac:

`/Users/tavo_vardas/511_invoice_tool`

### 3. Terminale nueik į folderį

Windows:

```bash
cd C:\511_invoice_tool
```

Mac:

```bash
cd /Users/tavo_vardas/511_invoice_tool
```

### 4. Įdiek bibliotekas

```bash
pip install -r requirements.txt
```

## Naudojimas per naršyklę

Paleisk:

```bash
streamlit run app.py
```

Atsidarys lokalus puslapis naršyklėje. Įkeli PDF, gauni Excel.

## Naudojimas per komandą

```bash
python converter_511_invoice.py SE.IN-00244072.pdf
```

Tame pačiame folderyje atsiras:

`SE.IN-00244072.xlsx`

## Svarbu

Jeigu PDF formatas pasikeis ir įrankis nebegalės patikimai nustatyti stulpelių ribų, jis Excel negeneruos.
Tai yra teisingas elgesys - geriau klaida nei sugadinti duomenys.


## Versija v2

Pataisyta lentelės ribų aptikimo logika:
- ignoruojamos antraštės vidinės linijos;
- naudojamos tik vertikalios ribos, kurios eina per prekių eilutes;
- tai neleidžia Description daliai patekti į Color stulpelį.


## Versija v3

Pataisyta trumpų paskutinių puslapių problema:
- jeigu paskutiniame puslapyje per mažai vertikalių linijų pakartojimų, naudojamos anksčiau patvirtintos lentelės ribos;
- eilutės aptinkamos pagal faktinį Line stulpelį ir nuoseklią eilučių seką;
- pridėta blokavimo kontrolė, kad Line seka būtų 1..paskutinė eilutė be tarpų.


## Versija v4

Pataisyta praleisto Line numerio problema:
- kai PDF realiai neturi vieno Line numerio, pvz. praleista 44 eilutė, įrankis nebestabdo tolesnių eilučių;
- tikrinama, kad Line numeriai būtų griežtai didėjantys ir be dublikatų;
- Amount suma privalo sutapti su INVOICE TOTAL, todėl trūkstamos eilutės vis tiek būtų pagautos.


## Versija v5

Pataisyta mažų sąskaitų problema:

- v4 per griežtai reikalavo, kad item-table vertikalios ribos kartotųsi daug kartų;
- mažose sąskaitose su 1–6 pozicijomis ribų kartojimų per mažai, todėl būdavo klaida:
  `Nepavyko ištraukti nė vienos prekių lentelės eilutės.`
- v5 palieka v4 griežtą aptikimą normalioms sąskaitoms;
- jeigu pirmame prekių puslapyje strict aptikimas nepavyksta, naudojamas mažų sąskaitų fallback;
- tęstiniuose/finaliniuose puslapiuose vis dar naudojamos anksčiau patvirtintos ribos;
- Customer PO dabar leidžia reikšmes su tarpais, pvz. `Special price`;
- Color / Size / Dim toliau ištraukiami tik pagal realias PDF x-koordinates, ne pagal tekstinį spėjimą.


## Versija v6

Pridėtas bulk upload:

- galima įkelti vieną arba kelias PDF sąskaitas vienu metu;
- kiekviena sąskaita konvertuojama atskirai;
- sėkmingi Excel failai rodomi atskirai ir gali būti atsisiunčiami po vieną;
- visi sėkmingi Excel failai gali būti atsisiunčiami vienu ZIP;
- ZIP viduje papildomai įdedamas `validation_report.txt`;
- nepavykusių PDF klaidos rodomos prie konkretaus failo;
- klaidų sąrašą galima atsisiųsti kaip `conversion_errors.txt`;
- konverterio logika liko v5: mažų sąskaitų fallback + Customer PO su tarpais.


## Versija v7

Pridėtas slaptažodis web versijai:

- jei `APP_PASSWORD` nėra nustatytas, lokalus įrankis veikia be slaptažodžio;
- Streamlit Cloud reikia įrašyti secret:
  `APP_PASSWORD = "tavo_slaptažodis"`
- slaptažodis nelaikomas kode;
- bulk upload funkcija iš v6 palikta;
- konverterio logika liko v5/v6: mažų sąskaitų fallback + Customer PO su tarpais.

### Streamlit Cloud Secrets

Streamlit Cloud:

1. Atidaryk app valdymą.
2. Eik į Settings / Secrets.
3. Įrašyk:

```toml
APP_PASSWORD = "tavo_sugalvotas_slaptazodis"
```

4. Save.
5. Reboot / rerun app.

### GitHub

Į GitHub kelti šiuos failus:

- `app.py`
- `converter_511_invoice.py`
- `requirements.txt`
- `README.md`

Nekelti PDF sąskaitų, Excel failų, ZIP rezultatų ir `__pycache__`.


## Versija v8

Pataisyta lokalaus paleidimo klaida:

- v7 lokaliai mesdavo `StreamlitSecretNotFoundError`, jeigu nebuvo `.streamlit/secrets.toml`;
- v8 lokaliai veikia be slaptažodžio, jeigu secrets failo nėra;
- Streamlit Cloud aplinkoje slaptažodis veikia per `APP_PASSWORD` secret.
