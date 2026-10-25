# Norröna SuperReport — præcise nøgletal

En enkel hjemmeside, der sammenholder passagerfordeling med verificerbare tal fra JetReports.

## Start

Åbn https://kronborg1980.github.io/Norrona-SuperReport-/ og indlæs én eller flere Excel-filer med drag-and-drop. Du kan indlæse pax-år for 2024, 2025, 2026 osv. samt regnskab for TFS, FNB, REC, ACC og GAL. Vælg år, periode, afdeling og bookingstatus.

Filer for samme passagerår erstatter tidligere udgaver af det år. En JetReport for samme år og afdeling erstatter den tidligere rapport. Ingen passager- eller regnskabstal ligger i kildekoden.

## Hvilke nøgletal er faktiske?

- Omsætning: Income total, Act.-kolonnen i JetReports.
- Officielle passagerer: Pax:, Act.-kolonnen i JetReports. Skibets pax tælles aldrig flere gange på grund af flere afdelinger.
- **Omsætning pr. pax:** Summen af omsætning divideret med summen af JetReports-pax i de samme måneder. Det er ikke gennemsnittet af månedernes Sales pr. Pax.
- Dækningsbidrag: Gross-Profit for én afdeling.
- Dækningsgrad: Samlet Gross-Profit divideret med samlet Income total × 100.
- Resultat: Total, Act.-kolonnen for én afdeling.
- Nationalitetsandel: Registreringer fra landet divideret med samtlige passagerregistreringer i samme periode og bookingstatusfilter.
- År-mod-år på økonomi vises kun, når der er JetReports for samme valgte måneder i begge år.
- Manglende måneder vises som en tankestreg, ikke som nul.
- Bookingstatus: Alle registreringer inklusive NSH som standard, med mulighed for at udelukke NSH.
- Ved valg af Alle indlæste afdelinger vises kun summen af de indlæste afdelingers omsætning. Ikke et fejlagtigt samlet skibsresultat.
- **Ingen simulation eller nationalitetsspecifik omsætning:** De to filtyper kan ikke vise, hvilke nationaliteter der faktisk har købt varerne.

## Kontrol mod Taxfree september 2026

| Periode | Omsætning | JetReports-pax | DKK/pax |
| --- | ---: | ---: | ---: |
| September 2026 | 2.337.598,36 DKK | 10.785 | 216,7453 |
| Januar–september 2026 | 30.161.795,09 DKK | 98.671 | 305,6804 |

September-tallet stemmer med Sales pr. Pax-rækken i den originale JetReport.

## Filer

- index.html: brugerflade og JetReports-inspireret design.
- app-v2.js: import, afstemning, rigtige nøgletal, tabeller og eksport.
- Excel behandles i browseren med SheetJS; PowerPoint genereres med PptxGenJS.

Ældre simuleringer, estimater og rapportafsnit er fjernet fra den aktuelle hjemmeside.
