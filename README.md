# Sprungmarken (SPFx 1.4.1, React 16, Fluent UI 8)

Port des SharePoint-internen `AnchorService` als **Application Customizer** plus **Inhaltsverzeichnis-Webpart**.
Läuft auch auf **SharePoint 2019**, wo es keine eingebauten Sprungmarken gibt.

| Komponente | Aufgabe |
|---|---|
| **AnchorService** (Application Customizer, `src/extensions/anchorService`) | Läuft auf jeder Seite der Website: vergibt IDs an Überschriften, Link-Symbol, wandelt `{anchor-#id}`-Marker um, springt zum `#hash` – auch nach Seitenwechsel und Bearbeiten |
| **Inhaltsverzeichnis** (Webpart, `src/webparts/tableOfContents`) | Zeigt die Anker des Services als Baum oder eigene Links (`PropertyFieldCollectionData` mit Anker-Dropdown) |

Beide liegen im Bundle `anchor-links`. Der gemeinsame Zustand (`anchorServiceManager`, Seitenwechsel,
Konfiguration, CSS-Klassen) liegt zusätzlich am `window` (`__spfxAnchorService_v1`), damit Extension und
Webpart sich auch in getrennten Bundles oder Versionen dieselben Anker teilen.

## Extension einrichten

**Mit der App (Feature):** Beim Hinzufügen der App zur Website hängt `sharepoint/assets/elements.xml` den
Customizer ein (`skipFeatureDeployment: false`).

**Per PnP PowerShell** (z. B. andere Eigenschaften je Website):

```powershell
Add-PnPCustomAction -Name "AnchorService" -Title "AnchorService" `
  -Location "ClientSideExtension.ApplicationCustomizer" `
  -ClientSideComponentId "38464b1c-f030-48cd-93ca-7b1be4f645d1" `
  -ClientSideComponentProperties '{"headingLevels":[2,3,4],"showAnchorIcons":true,"convertInlineAnchors":true,"scrollToHashOnLoad":true}'
```

| Eigenschaft | Standard | Bedeutung |
|---|---|---|
| `headingLevels` | `[1,2,3,4,5]` | Verankerte Ebenen (Array oder Text `"2,3,4"`) |
| `showAnchorIcons` | `true` | Link-Symbol neben Überschriften |
| `convertInlineAnchors` | `true` | `{anchor-#id}` (auch `{Anker-#…}`, `{ANCHOR-#…}`) in Absätzen → unsichtbares Sprungziel |
| `scrollToHashOnLoad` | `true` | Beim Laden/Seitenwechsel zum `#hash` springen |
| `contentSelector` | leer | CSS-Selektor des Inhaltsbereichs, leer = automatisch |

Debuggen: `gulp serve --config anchorService` (siehe `config/serve.json`, `pageUrl` anpassen).

## Inline-Anker `{anchor-#id}`

Inline-Anker sind vollwertige Anker des AnchorService (`kind: 'inline'` in der Ankerliste):

- **Überall im Inhalt:** Absätze, Listen, Tabellen, auch in Überschriften (TreeWalker über den Inhaltsbereich,
  kein eigener `p`-Selektor mehr). Ausgenommen: `<pre>`, `<code>`, Formularfelder, Skripte.
- **Feste ID:** Die ID bleibt exakt wie geschrieben (z. B. `Wichtig_1`) und wird **vor** den Überschriften reserviert.
  Heißt eine Überschrift gleich, weicht die Überschrift aus (`einleitung-1`) – Links auf den Marker brechen nie.
  Doppelte Marker: der zweite bekommt ein Suffix (`Doppelt-1`).
- **Eine Logik für alles:** `collectAnchorCandidates` (sammeln) und `assignAnchorIds` (IDs vergeben) werden von der
  Extension (Ansicht) und von der Vorhersage (`getAvailableAnchors` / `usePredictedAnchors`: Vorschau und Dropdown im
  Bearbeitungsmodus) gemeinsam genutzt – vorhergesagte und echte IDs sind dadurch identisch.
- Scrollen zu Inline-Ankern läuft über den Service (`anchorServiceManager.scrollTo('Wichtig_1')`).

## Bearbeitungsmodus

Eine Extension kennt keinen `displayMode`. Erkannt wird Bearbeiten über:
1. das Inhaltsverzeichnis-Webpart meldet seinen `displayMode` (`anchorEngine.reportEditMode`),
2. `?Mode=Edit` in der URL,
3. einen bearbeitbaren Editor (`contenteditable`) im Canvas,
4. Bearbeitungs-Klassen des Canvas.

Im Bearbeitungsmodus werden Icons/Wrapper entfernt und `{anchor-#…}`-Marker wiederhergestellt.
**Schutz:** Elemente in einem `contenteditable`-Bereich werden nie verändert – auch nicht, falls der Modus zu spät erkannt wird.
`?Mode=Edit` gilt nicht als Seitenwechsel.

## Seitenwechsel

Die Extension bleibt bei der clientseitigen Navigation bestehen. Erkennung über `application.navigatedEvent`,
`popstate` und eine URL-Prüfung bei DOM-Änderungen. Dann: alte Seite aufräumen → `anchorServiceManager.reset()` →
neue Seite verankern → Hash der neuen URL anspringen. Ein neu gerenderter Canvas (z. B. nach Veröffentlichen)
wird ebenfalls erkannt.

## Inhaltsverzeichnis

- Quelle „Überschriften der Seite“: angezeigt werden die im Webpart gewählten Ebenen, **soweit die Extension sie verankert**
  (der Eigenschaftenbereich zeigt die verankerten Ebenen an). Im Bearbeitungsmodus zeigt die Vorschau bereits die echten Ziel-IDs.
- Quelle „Eigene Links“: `PropertyFieldCollectionData` mit Titel (`LinkDropdown`), URL und Ebene. Das Dropdown sagt die
  IDs mit der Konfiguration der Extension voraus (gleiche Ebenen, Filter, Inhaltsbereich; Inline-Anker nur wenn umgewandelt).
- Filter: `<aside>`, `data-toc-ignore`, `display: none`, Site-Header-Titel, außerhalb `#spPageChromeAppDiv`.
- Ist die Extension nicht aktiv, zeigt das Webpart im Bearbeitungsmodus einen Hinweis.

### Validierung eigener Links (`src/anchors/linkValidation.ts`)

Grundsatz: Alles, was auf einer Seite als Sprungmarke entstehen kann, ist gültig.

| Feld | Gültig | Abgelehnt |
|---|---|---|
| Titel | 1–255 Zeichen, jedes druckbare Zeichen (é è ê œ « » „ “ – € 中文 🚀 …) | leer, > 255, Steuerzeichen |
| URL | `#id` roh oder kodiert (`#été-–-résumé`, `#a.1-grundlagen`, `#Wichtig_1`), relative/serverrelative Adressen (auch mit Leerzeichen), `//host`, `http(s):`, `mailto:`, `tel:` | leer, > 2048, `< > " \``, Steuerzeichen, `javascript:`/`data:`/`vbscript:`/`file:` …, leeres `#`, Anker mit Leerzeichen oder `# { } \| \ ^`, kaputte %-Kodierung |

## Sprachen

Deutsch, Englisch, Französisch (Webpart, Extension, Manifest).

## Bibliothek `src/anchors`

| Datei | Inhalt |
|---|---|
| `AnchorServiceManager.ts` | `anchorServiceManager` (inkl. `reset()`), `AnchorService`, `allocateAnchorId` |
| `PageAnchors.tsx` | Engine-Komponente (ohne Anzeige), von der Extension gehostet |
| `anchorEngine.ts` | Konfiguration der Extension, Bearbeitungsmodus-Erkennung, `useAnchorEngine` |
| `pageNavigation.ts` | Seitenwechsel-Hub, `navigatedEvent`, Zonenbesitz |
| `globalRegistry.ts` | seitenweite Singletons über Bundle-Grenzen |
| `AnchorZone.tsx` / `HiddenAnchorZone.tsx` | Sprungmarken für Überschriften / eingeklappte Bereiche |
| `anchorIds.ts` | ID-Vergabe: `sanitizeAnchorId`, Kollisionen, `assignAnchorIds` (feste Inline-IDs zuerst) |
| `anchorCollector.ts` | `collectAnchorCandidates`: Überschriften + Inline-Anker in Dokumentreihenfolge |
| `inlineAnchors.ts` | `MARKER_REGEX`, `getQuerySelector`, `getInlineAnchorsElements`, `convertInlineAnchors`, `revertInlineAnchors`, `stripInlineMarkers` |
| `anchorCatalog.ts` | Sprungmarken der Seite vorhersagen (Dropdown) |
| `linkValidation.ts` | Validierung von Titel und URL |
| `elementFilters.ts`, `contentRoot.ts` | Filter, Inhaltsbereich |
| `hooks.ts` | `useSortedAnchors` (Ansicht), `usePredictedAnchors` (Bearbeiten), `usePageVersion`, `useDomChanges` |
| `anchorStyles.ts`, `pageTheme.ts` | Styles (`mergeStyleSets`), Seiten-Theme |

## Build

```bash
npm install
gulp bundle --ship && gulp package-solution --ship   # -> sharepoint/solution/spfx-anchor-links.sppkg
```

Geprüft mit TypeScript 4.9 gegen React 16.14, `@fluentui/react` 8, `@pnp/spfx-property-controls` 3 und den
SPFx-1.4.1-Paketen. Versionen in `package.json` an euer Projekt angleichen.
