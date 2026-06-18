# Sticker Generator

Webpagina voor het maken van stickers vanuit een Halindeling .xlsx bestand.

## Wat doet het?

1. Je uploadt de halindeling
2. Je kiest welke locatie-prefixen je wilt (eerste 2 tekens, bv. `gK`, `gL`, `bA`, `eT`)
3. Optioneel: je filtert op klantcode-prefix
   - Codes met een **cijfer** aan het begin → eerste **3** tekens (bv. `6GF155` → `6GF`)
   - Codes met een **letter** aan het begin → eerste **2** tekens (bv. `EF168` → `EF`)
4. Je downloadt een PDF met 1 sticker per unieke klant (10×15 cm, locatie groot boven, klantcode kleiner onder, 90° gedraaid)

## Installeren

In de map `sticker-app/`:

```
npm install
```

(Vereist [Node.js](https://nodejs.org/) — versie 16 of hoger.)

## Starten

```
npm start
```

Open daarna in je browser: <http://localhost:3000>

## Aanpassen

- Locaties met een leading `g` worden gestript op de sticker (bv. `gK3 1` wordt geprint als `K3 1`).
- Sticker formaat (10×15 cm), verhouding locatie/klant (4:1) en marges zijn instelbaar via constanten bovenin `server.js`.

## Bestanden

- `server.js` — Express backend, parsing en PDF generatie
- `public/index.html` — frontend (upload + filters)
- `package.json` — afhankelijkheden
