# Norröna SuperReport

Selvstændig analyse af passagermix og kommercielle JetReports for MS Norröna.

## Start

Åbn `super-report/index.html` gennem en webserver / dit eksisterende JetReports-hosting. Vælg **Indlæs Excel**, eller træk filerne direkte ind på siden.

Indlæs:
- Passagerfiler: `Catrgory pr departure 2025.xlsx`, `... 2026.xlsx` og senere 2024/2027.
- JetReports: de aktuelle `Mánaðartøl`-udtræk for **TFS, FNB, REC, ACC og GAL**, et sæt for hvert regnskabsår.

Filer kan indlæses i vilkårlig rækkefølge. Et nyt udtræk for samme passagerår erstatter det forrige år. Et nyt JetReport-udtræk for samme år og afdeling erstatter den tidligere version. Filerne gemmes ikke som data i projektet.

## Analyse

- Nationaliteter, landekoder, no-shows, års- og månedssammenligninger
- Omsætning, dækningsbidrag og resultat pr. afdeling
- Omsætning pr. JetReports-passager
- Månedlig samvariation mellem andelen af en valgt nationalitet og omsætning pr. passager; samvariation betyder **ikke** årsag
- Samme-kalendermåned år-for-år hvor der findes tilstrækkeligt mange datapunkter
- Volumen- vs. forbrugseffekt ved årssammenligning
- Afstemning mellem passagerexport og JetReports-pax
- Ledelsesrapport, CSV, print/PDF og PowerPoint med Smyril Line/JET-lignende grafisk tema

## Centrale beregningsprincipper

1. **Alle afdelinger kræver alle fem afdelingsrapporter for hver måned.** Delvise skibstotaler vises ikke som fuldstændige.
2. **Pax fra JetReports må ikke lægges sammen over afdelinger.** De fem rapporter viser skibets samme passagertal.
3. **Kabyssens vareforbrug:** F&B's interne galley allocation tilbageføres ved samlet FNB+GAL-resultat/DB, fordi GAL allerede viser selve omkostningen.
4. **Nationalitetsudgifter/omsætning:** De vedhæftede tal indeholder ingen entydig transaktionskobling mellem kunde og betaling. SuperReport beregner derfor ikke omsætning per tysk, dansk eller færøsk gæst. Det analyserer i stedet ændringer i mix og månedsomsætning pr. passager.
5. **No-shows:** Udelukkes som standard fra passagerfilens registreringer. JetReports' egne pax bruges til nøgletallet omsætning pr. passager.
6. **Samme perioder:** Årsændringer vises kun for samme dækkede måneder, og manglende regnskabsdata behandles ikke som nul.

## Drift

En enkelt statisk HTML-fil. Excel-filerne indlæses af SheetJS i browseren. PowerPoint genereres af PptxGenJS, der indlæses, når brugeren vælger eksport. Bibliotekerne hentes fra CDN, så internetforbindelse er nødvendig første gang.


## Passagermix-dybde (version 2)

- **All pax som standard:** Alle nationaliteter og alle bookingstatusser/registreringer vises ved start. Bookingstatus-filteret kan skiftes til uden no-show eller kun no-show.
- **Mix-dybde:** Opdel antal og andel efter land, nationalitetsgruppe, rute, passagerkategori, alder, bookingstatus eller rækketype. Ekstra filtre for land, rute, kategori og alder.
- **År-mod-år:** Se forskel i antal, procent og procentpoint for de samme måneder.
- **Krydstabeller:** Nationalitet × rute og nationalitet × passagerkategori, plus stigninger, fald og månedlig andelsudvikling.
- **Sammenhæng:** Landekoders andel over for omsætning pr. JetReports-passager i den valgte afdeling. Pearson r beregnes kun ved mindst seks sammenhørende måneder med tilstrækkelig variation. Samme kalendermåned år-mod-år vises særskilt.
- **Dybere rapport og PowerPoint:** Nationalitetsforskydninger, ruter, kategorier, alder og no-shows. Ni slides i den eksisterende Smyril Line-inspirerede stil.
- **Passagerfiler alene:** Fælles måneder og passagermix kan analyseres uden indlæste JetReports. Omsætning kræver de tilsvarende regnskaber.

**Definition:** “All pax” betyder alle registreringer i passagereksporten, inklusive NSH. Det er ikke nødvendigvis identisk med JetReports' officielle rejsetal. SuperReport viser afvigelsen under Datakontrol og anvender JetReports-pax som nævner for omsætning pr. passager.
