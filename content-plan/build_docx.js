const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun,
  HeadingLevel, PageBreak, ShadingType, BorderStyle, ExternalHyperlink, AlignmentType } = require('docx');

const EX = path.join(__dirname, 'examples');
const days = [
 {h:'Poniedziałek · 2 listopada', a:['12:00 · Karuzela „Wracamy z targów”',
  'Za nami [nazwa targów] — dziękujemy, że byliście z nami! Przez kilka dni rozmawialiśmy z projektantami i architektami. W karuzeli najlepsze chwile ze stoiska. Nie zdążyliście nas odwiedzić? Napiszcie — pokażemy wszystkie modele w showroomie.',
  '6 slajdów wygenerowanych przez AI z istniejących zdjęć stoiska i mebli: widok ogólny, łóżko, tkaniny z bliska. Bez ludzi.',
  'Napisz w wiadomości','#infiore #mebletapicerowane #meblenazamowienie #projektowaniewnętrz',
  '@warsawhome_official','https://www.instagram.com/p/De9FZxdmucG/','2026-11-02_1200_warsawhome_sm.jpg'],
  b:['19:30 · Reels „Nasze stoisko w 30 sekund”',
  'Tak wyglądało nasze stoisko na [nazwa targów]! Łóżka, banketki i szafki nocne — wszystko na zamówienie z naszej pracowni. Który model podobał się Wam najbardziej? Napiszcie w komentarzu.',
  'Wideo 20–30 s wygenerowane przez AI z istniejących zdjęć: płynny przejazd kamery, zbliżenia przeszyć. Bez ludzi.',
  'Zostaw komentarz','#infiore #łóżkotapicerowane #sypialnia #inspiracjesypialnia',
  '@meblokos','https://www.instagram.com/reel/DfCCM-lBWxW/','2026-11-02_1930_meblokos_sm.jpg'],
  l:'Logika: w południe karuzela zbiera zapisy i wiadomości, wieczorem rolka daje zasięg. Do uzupełnienia: nazwa targów.'},
 {h:'Wtorek · 3 listopada', a:['12:00 · Zdjęcie „Gwiazda stoiska”',
  'O ten model pytaliście najczęściej — [nazwa modelu]. Miękki, wysoki zagłówek, czyste linie i tkanina, którą chce się dotknąć. Wymiar, kolor i materiał dobierzemy do Twojego wnętrza. Chcesz poznać cenę w swoim rozmiarze? Napisz do nas.',
  'Jedno zdjęcie łóżka z targów, dopracowane przez AI: aranżacja sypialni, światło dzienne, format 4:5. Bez ludzi.',
  'Zapytaj o cenę','#infiore #łóżkotapicerowane #sypialnia #meblenazamowienie',
  '@dappi_mebletapicerowane','https://www.instagram.com/p/DervTphEUW-/','2026-11-03_1200_dappi_sm.jpg'],
  b:['19:30 · Reels „Od szkicu do stoiska”',
  'Zanim stoisko zachwyciło gości, były tygodnie pracy: szkice, dobór tkanin, szycie, tapicerowanie, montaż. Tak powstaje każde Infiore — powoli i z dbałością o detal.',
  'Wideo 15–20 s wygenerowane przez AI z istniejących zdjęć: szkic → próbki tkanin → przeszycia z bliska → gotowe łóżko. Bez ludzi i głosu, tylko muzyka.',
  'Obserwuj nas','#infiore #rzemiosło #tapicerstwo #handmadeinpoland',
  '@softwall.eu','https://www.instagram.com/reel/DZjuh_mMKVT/','2026-11-03_1930_softwall_sm.jpg'],
  l:'Logika: w południe mocne zdjęcie produktu prowadzi do zapytań o cenę, wieczorem klimatyczne wideo pokazuje rzemiosło. Do uzupełnienia: nazwa modelu.'},
 {h:'Środa · 4 listopada', a:['12:00 · Karuzela „Pytania projektantów z targów”',
  'Zebraliśmy pytania, które najczęściej zadawali nam projektanci na [nazwa targów]: niestandardowy wymiar? własna tkanina? czas realizacji? rysunek techniczny? Odpowiedzi w karuzeli. Projektujesz wnętrze? Napisz — przygotujemy ofertę.',
  '5–6 slajdów: duże pytanie + krótka odpowiedź na tle zdjęć mebli dopracowanych przez AI. Bez ludzi.',
  'Napisz w wiadomości','#infiore #projektowaniewnętrz #architekturawnętrz #meblenazamowienie',
  '@maniaspania','https://www.instagram.com/p/De34WvIDqtC/','2026-11-04_1200_maniaspania_sm.jpg'],
  b:['19:30 · Reels „Goście o Infiore”',
  'Co mówili o nas goście targów? Kilka krótkich opinii, które zapamiętaliśmy najbardziej. Dziękujemy za każdą rozmowę — to dla nas najlepsza inspiracja.',
  'Wideo 15–20 s z AI: opinie gości jako napisy na tle ujęć łóżek i tkanin. Bez ludzi i głosu, tylko muzyka.',
  'Skomentuj','#infiore #mebletapicerowane #sypialnia #opinie',
  '@maniaspania','https://www.instagram.com/reel/DfCjliMu6gN/','2026-11-04_1930_maniaspania_sm.jpg'],
  l:'Logika: w południe treść dla projektantów (B2B), wieczorem społeczny dowód. Do uzupełnienia: czas realizacji, cytaty gości.'},
 {h:'Czwartek · 5 listopada', a:['12:00 · Zdjęcie „In fiore — w rozkwicie”',
  'Skąd nazwa Infiore? Po włosku „in fiore” znaczy „w rozkwicie”. Tak chcemy, żeby czuły się nasze wnętrza — żywe, ciepłe i pełne charakteru. Każdy mebel tworzymy na zamówienie, żeby rozkwitał razem z Twoim domem.',
  'Zdjęcie łóżka z kwiatami na szafce nocnej, dopracowane przez AI: ciepłe światło, format 4:5. Bez ludzi.',
  'Obserwuj nas','#infiore #sypialnia #wnętrza #madeinpoland',
  '@meblesid.pl','https://www.instagram.com/p/C_0G9M8sM31/','2026-11-05_1200_meblesid_sm.jpg'],
  b:['19:30 · Reels „Detale, które robią różnicę”',
  'Równe przeszycia, idealnie naciągnięta tkanina, starannie wykończone narożniki. To właśnie odróżnia mebel robiony ręcznie.',
  'Wideo 15 s z AI: makro-ujęcia przeszyć, pikowania, nóżek i krawędzi, powolne ruchy kamery. Bez ludzi.',
  'Zapisz post','#infiore #rzemiosło #tapicerstwo #detale',
  '@dappi_mebletapicerowane','https://www.instagram.com/reel/DfFldLRx0r-/','2026-11-05_1930_dappi_sm.jpg'],
  l:'Logika: w południe historia marki, wieczorem jakość wykonania.'},
 {h:'Piątek · 6 listopada', a:['12:00 · Karuzela „Banketka — mały mebel, duża zmiana”',
  'Banketka to jeden z najbardziej uniwersalnych mebli. W nogach łóżka, w przedpokoju, w garderobie czy przy toaletce — wszędzie dodaje wygody i elegancji. W karuzeli 4 miejsca, w których sprawdzi się najlepiej. Którą wersję wybierasz?',
  '5 slajdów z AI: ta sama banketka w sypialni, przedpokoju, garderobie, przy toaletce + slajd z tkaninami. Bez ludzi.',
  'Skomentuj','#infiore #banketka #meblenazamowienie #sypialnia',
  '@maniaspania','https://www.instagram.com/p/Dd1kGPkiUjd/','2026-11-06_1200_maniaspania_sm.jpg'],
  b:['19:30 · Reels „Nowości z targów w showroomie”',
  'Modele z targów są już w naszym showroomie. Przyjdź, usiądź, dotknij tkanin i porozmawiaj z nami o swoim wnętrzu. [adres showroomu]. Umów wizytę w wiadomości.',
  'Wideo 20 s z AI z istniejących zdjęć showroomu: przejście między meblami, na końcu adres. Bez ludzi.',
  'Umów wizytę','#infiore #showroom #mebletapicerowane #warszawa',
  '@maniaspania','https://www.instagram.com/reel/Ddt0rj2FCpU/','2026-11-06_1930_maniaspania_sm.jpg'],
  l:'Logika: w południe produkt i pomysły, wieczorem zaproszenie na weekend do showroomu. Do uzupełnienia: adres showroomu.'},
 {h:'Sobota · 7 listopada', a:['12:00 · Zdjęcie „Weekendowy poranek”',
  'Sobotni poranek bez budzika. Miękki zagłówek, ulubiona pościel i kawa w łóżku. Czasem tyle wystarczy, żeby poczuć się jak w domu. Miłego weekendu!',
  'Zdjęcie łóżka z pościelą i tacą z kawą, wygenerowane przez AI z istniejącego zdjęcia łóżka, poranne światło. Bez ludzi.',
  'Zapisz post','#infiore #sypialnia #weekend #inspiracjesypialnia',
  '@maniaspania','https://www.instagram.com/p/DdosP-VjczX/','2026-11-07_1200_maniaspania_sm.jpg'],
  b:['19:30 · Karuzela „Wybierz tkaninę”',
  'Ten sam model, trzy charaktery. Który wybierasz? A — welur w kolorze butelkowej zieleni, B — jasne boucle, C — naturalny len. Napisz literę w komentarzu!',
  '3 slajdy z AI: jedno łóżko Infiore w trzech tkaninach, podpisane A, B, C. Bez ludzi.',
  'Skomentuj','#infiore #tkaniny #łóżkotapicerowane #aranżacjawnętrz',
  '@maniaspania','https://www.instagram.com/reel/Ddjn-bvDFNW/','2026-11-07_1930_maniaspania_sm.jpg'],
  l:'Logika: w weekend lżejsze treści, wieczorem zabawa w wybór, która zbiera komentarze.'},
 {h:'Niedziela · 8 listopada', a:['12:00 · Zdjęcie „Niedziela w sypialni”',
  'Niedziela to idealny dzień, żeby zostać dłużej w łóżku. Jeśli sypialnia ma być miejscem odpoczynku, warto zadbać o każdy jej element — od zagłówka po szafkę nocną.',
  'Spokojna aranżacja sypialni z łóżkiem i szafką nocną, wygenerowana przez AI z istniejących zdjęć, stonowane kolory. Bez ludzi.',
  'Zapisz post','#infiore #sypialnia #niedziela #wnętrza',
  '@maniaspania','https://www.instagram.com/p/Ddb7AyaDIqi/','2026-11-08_1200_maniaspania_sm.jpg'],
  b:['19:30 · Reels „Dziękujemy za ten tydzień”',
  'To był intensywny tydzień po targach. Dziękujemy za wszystkie wiadomości, pytania i odwiedziny. W przyszłym tygodniu pokażemy, jak powstaje łóżko na zamówienie — krok po kroku.',
  'Wideo 15–20 s z AI: montaż najlepszych ujęć tygodnia, na końcu zapowiedź kolejnego. Bez ludzi.',
  'Obserwuj nas','#infiore #mebletapicerowane #meblenazamowienie #madeinpoland',
  '@maniaspania','https://www.instagram.com/p/Det_AyZDsH2/','2026-11-08_1930_maniaspania_sm.jpg'],
  l:'Logika: spokojne zamknięcie tygodnia i zapowiedź następnego.'},
];

const W = [1500, 3763, 3763]; // A4 width 11906 - 2*1440 margins = 9026
const FONT = 'Arial';
const border = { style: BorderStyle.SINGLE, size: 4, color: 'C9C6BF' };
const borders = { top: border, bottom: border, left: border, right: border };
const t = (s, o = {}) => new TextRun({ text: s, font: FONT, size: 19, ...o });
const cell = (children, i, o = {}) => new TableCell({ width: { size: W[i], type: WidthType.DXA }, borders,
  margins: { top: 70, bottom: 70, left: 100, right: 100 }, children, ...o });
const P = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs], spacing: { after: 40 }, ...o });

function img(file) {
  const buf = fs.readFileSync(path.join(EX, file));
  // read jpeg dims
  let i = 2, w = 200, h = 180;
  while (i < buf.length) { if (buf[i] !== 0xFF) break; const m = buf[i + 1]; const len = buf.readUInt16BE(i + 2);
    if (m >= 0xC0 && m <= 0xC3) { h = buf.readUInt16BE(i + 5); w = buf.readUInt16BE(i + 7); break; } i += 2 + len; }
  const tw = 150; return new ImageRun({ type: 'jpg', data: buf, transformation: { width: tw, height: Math.round(h * tw / w) } });
}
const row = (label, a, b, shade) => new TableRow({ children: [
  cell([P(t(label, { bold: true }))], 0, shade ? { shading: { fill: 'EFEDE8', type: ShadingType.CLEAR } } : {}),
  cell(a, 1, shade ? { shading: { fill: 'EFEDE8', type: ShadingType.CLEAR } } : {}),
  cell(b, 2, shade ? { shading: { fill: 'EFEDE8', type: ShadingType.CLEAR } } : {})] });
const link = (label, url) => P([new ExternalHyperlink({ link: url, children: [t(label, { color: '2F4A3E', underline: {} })] })]);

const children = [];
days.forEach((d, idx) => {
  if (idx > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
  if (idx === 0) {
    children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun({ text: 'Infiore · plan treści · 2–8 listopada 2026', font: FONT, size: 32, bold: true })], spacing: { after: 120 } }));
  }
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: d.h, font: FONT, size: 28, bold: true, color: '2F4A3E' })], spacing: { after: 140 } }));
  const A = d.a, B = d.b;
  children.push(new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: W, rows: [
    row('', [P(t(A[0], { bold: true }))], [P(t(B[0], { bold: true }))], true),
    row('Tekst', [P(t(A[1]))], [P(t(B[1]))]),
    row('Wizual (AI)', [P(t(A[2]))], [P(t(B[2]))]),
    row('Wezwanie', [P(t(A[3]))], [P(t(B[3]))]),
    row('Hashtagi', [P(t(A[4], { color: '2F4A3E' }))], [P(t(B[4], { color: '2F4A3E' }))]),
    row('Przykład', [link(A[5], A[6]), P([img(A[7])])], [link(B[5], B[6]), P([img(B[7])])]),
  ] }));
  children.push(new Paragraph({ children: [t(d.l, { italics: true })], spacing: { before: 160 } }));
});

const doc = new Document({ styles: { default: { document: { run: { font: FONT } } } },
  sections: [{ properties: { page: { margin: { top: 1134, bottom: 1134, left: 1440, right: 1440 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(path.join(__dirname, 'Infiore_plan_2-8_listopada_2026.docx'), b); console.log('ok'); });
