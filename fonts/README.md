# Bundled fonts

Open fonts that stand in for the common Office faces when a Word file is laid out
(`document-kit/layout`). Each family is metric-compatible with the face it stands in for: the
same advance widths, so lines break where Word breaks them. Which document font goes to which
family is the engine's own table (`resolveFamilyKey`, `src/ream/core/fonts/remote-fonts.ts`);
`standInFor(name)` in `document-kit/layout` reads it.

| Family | Stands in for | Version | Licence | Licence text |
|---|---|---|---|---|
| Tinos | Times New Roman; other serif faces | 1.23 | Apache-2.0 | `LICENSE-Apache-2.0.txt` |
| Arimo | Arial, Helvetica; the default for any unknown face | 1.33 | SIL OFL 1.1 | `LICENSE-OFL-Arimo.txt` |
| Carlito | Calibri, Calibri Light | 1.104 | SIL OFL 1.1 (Reserved Font Name "Carlito") | `LICENSE-OFL-Carlito.txt` |
| Caladea | Cambria | 1.001 | SIL OFL 1.1 | `LICENSE-OFL-Caladea.txt` |
| Cousine | Courier New; other monospaced faces | 1.21 | Apache-2.0 | `LICENSE-Apache-2.0.txt` |

Copyright: Tinos and Cousine "Digitized data copyright (c) 2010-2012 Google Corporation"
(Tinos and Cousine are trademarks of Google Inc.); Arimo "Copyright 2020 The Arimo Project
Authors (https://github.com/googlefonts/arimo)"; Carlito "Copyright 2013 The Carlito Project
Authors (https://github.com/googlefonts/carlito)"; Caladea "Copyright 2012 The Caladea Project
Authors (https://github.com/huertatipografica/Caladea)". The licence and copyright of each file
were read from its own `name` table (IDs 0, 13 and 14), which is what they are distributed
under. Arimo, often listed as Apache-2.0, carries the OFL in this version.

## Files

Four faces per family: `<Family>-Regular.ttf`, `-Bold.ttf`, `-Italic.ttf`, `-BoldItalic.ttf`,
renamed from the source packages' `<Family>_400Regular.ttf`, `_700Bold.ttf`,
`_400Regular_Italic.ttf`, `_700Bold_Italic.ttf`. The bytes are unchanged.

## Where they came from

The `@expo-google-fonts` npm packages on jsDelivr, at the addresses the engine's remote provider
uses, which jsDelivr served from these versions (18 November 2026):

- `@expo-google-fonts/tinos@0.2.3`, `@expo-google-fonts/arimo@0.2.3`,
  `@expo-google-fonts/caladea@0.2.3`, `@expo-google-fonts/cousine@0.2.3`: flat files,
  `https://cdn.jsdelivr.net/npm/@expo-google-fonts/<family>@0.2.3/<Family>_<face>.ttf`.
- `@expo-google-fonts/carlito@0.4.1`: nested files,
  `https://cdn.jsdelivr.net/npm/@expo-google-fonts/carlito@0.4.1/<face>/Carlito_<face>.ttf`.

The OFL texts are the Carlito package's `LICENSE_FONT` with each family's copyright line; the
Apache-2.0 text is <https://www.apache.org/licenses/LICENSE-2.0.txt>.

SHA-256:

```
054f3f97a338070c3558a9ce7e302565be85c7436b10bb0a01cf6a41f0782e61  Tinos-Regular.ttf
971a3f3e6a5cb3c0973074cd9384bbea2db3d647ee54240a8a749bb43e3569e7  Tinos-Bold.ttf
80f98dbb589cd1ede7585cb1863335e9dd149e000a152dbf86f8a2325fb52b76  Tinos-Italic.ttf
f2f805a8dd53f97ddd680c79a43c48cea8be26e3fdec5be354c23e61942b6c9c  Tinos-BoldItalic.ttf
02219a6ff8456d985e2f17f60bd5555bcfb29b47c13e8bccc8362447287b74c6  Arimo-Regular.ttf
d6a22848ffa3c9d57aff3b3419509f0728ceafa830b71d065f0e388b678aaf1f  Arimo-Bold.ttf
daf5c64af223c0be94cadb394225c27b5dd74aca31620af5e6cdf1b6b5ba9e27  Arimo-Italic.ttf
d2035458124d168366044ff3f7f59e26d00bf25187f9ddf8c2624e44eeb49906  Arimo-BoldItalic.ttf
ca019755404c45627a8566915df99068949dc32ee2bce48d6aeee7542d2a0a89  Carlito-Regular.ttf
51edbfa32d8af939913ae1f4ad0a5173e32083499218c133384638090295f0b0  Carlito-Bold.ttf
074cd1b89d53765d90d0ed3b4bfe49523efaaf4f3f430c006bc3233778b0ebb5  Carlito-Italic.ttf
25f5672c1985d168d6bc2973864fc5a7e374bb95fe8d0f91cff47ae17fa67691  Carlito-BoldItalic.ttf
951ac8213baf87771336e6fd5e7003c28f21a30cfa6b1ec842b02e1f380e47db  Caladea-Regular.ttf
df3d745453c019e0821bfba446fb28b77e894ce80e812012b0aaa03dbabc8f91  Caladea-Bold.ttf
2bec3661078d53f0e2570946b42f828cbacca2baf5e7b379527a3766d9043529  Caladea-Italic.ttf
877659e6267502e8e3e8af9ca36dfd9492e2c21c96edbec0995a6ebc45e1dda3  Caladea-BoldItalic.ttf
dcd526004fcfec4ec23a21ca458ea6111e1e0d4471c0ade2704927101b5dcedb  Cousine-Regular.ttf
30ce88d0c9d9834cde36018fa6332f87a11adf17eb48239cafebad9e21637344  Cousine-Bold.ttf
dc39e44bb6ee50ea8445210ce7871f3af836632024e84eed125785d1b3f24b05  Cousine-Italic.ttf
bc3f577ce4f0036c9f8045bfcb2f36ce8795de91915ccf127371d9e722d5e518  Cousine-BoldItalic.ttf
```

Not bundled: Noto Serif (the engine's source for Greek and Cyrillic in Cambria text) and the
Noto faces for non-Latin scripts. A document that needs them is laid out without them and the
gap is reported in `losses`.
