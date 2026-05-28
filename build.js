const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Travel Guide";
pres.title = "Путешествие по Барселоне";

// Design tokens
const cream = "E8DCC4";
const paper = "F4EBD8";
const terra = "C44536";
const blue  = "2A4A5C";
const ochre = "E1A140";
const ink   = "1A1A1A";
const muted = "8C7A5B";
const line  = "C8B68E";

const SLIDE_W = 13.333;
const TOTAL = 9;

// Shadow factory — never reuse the returned object (pptxgenjs mutates in-place)
const shadow = (opacity = 0.18) => ({
  type: "outer", color: "000000", blur: 12, offset: 4, angle: 90, opacity,
});

function addMosaicChips(slide, x, y, w, h, count, palette) {
  for (let i = 0; i < count; i++) {
    const size = 0.12 + Math.random() * 0.18;
    const cx = x + Math.random() * (w - size);
    const cy = y + Math.random() * (h - size);
    const color = palette[Math.floor(Math.random() * palette.length)];
    const rotate = Math.floor(Math.random() * 360);
    slide.addShape(pres.ShapeType.rect, {
      x: cx, y: cy, w: size, h: size,
      fill: { color },
      line: { type: "none" },
      rotate,
    });
  }
}

function addPageMark(slide, num, total) {
  const nn = String(num).padStart(2, "0");
  const tt = String(total).padStart(2, "0");
  slide.addText(`${nn} / ${tt}`, {
    x: 12.3, y: 7.05, w: 0.9, h: 0.3,
    fontFace: "Inter", fontSize: 9, color: muted,
    align: "right", charSpacing: 4,
  });
  slide.addText("BARCELONA", {
    x: 0.5, y: 7.05, w: 3, h: 0.3,
    fontFace: "Inter", fontSize: 9, color: muted,
    align: "left", charSpacing: 8, bold: true,
  });
}

function addEyebrow(slide, text, color = terra) {
  slide.addText(text, {
    x: 0.6, y: 0.5, w: 6, h: 0.3,
    fontFace: "Inter", fontSize: 10, color,
    charSpacing: 12, bold: true,
  });
}

function addPhotoBox(slide, x, y, w, h, caption) {
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h,
    fill: { color: cream },
    line: { color: line, pt: 0.75 },
    shadow: shadow(0.12),
  });
  slide.addShape(pres.ShapeType.line, {
    x, y, w, h,
    line: { color: line, pt: 0.5, dashType: "dash" },
  });
  slide.addShape(pres.ShapeType.line, {
    x, y: y + h, w, h: -h,
    line: { color: line, pt: 0.5, dashType: "dash" },
  });
  slide.addText("ФОТО", {
    x, y: y + h / 2 - 0.5, w, h: 0.6,
    fontFace: "Playfair Display", fontSize: 22, color: terra,
    italic: true, bold: true, align: "center",
  });
  slide.addText(caption, {
    x, y: y + h / 2, w, h: 0.45,
    fontFace: "Inter", fontSize: 10, color: muted,
    italic: true, align: "center",
  });
}

function addSpotSlide({ num, total, name, ruName, district, body, facts, photoCaption, photoOnLeft }) {
  const s = pres.addSlide();
  s.background = { color: paper };

  const photoX = photoOnLeft ? 0 : 7.0;
  addPhotoBox(s, photoX, 0.6, 6.5, 6.0, photoCaption);

  const cx = photoOnLeft ? 7.0 : 0.6;
  const cw = 5.7;

  s.addText(district.toUpperCase(), {
    x: cx, y: 0.6, w: cw, h: 0.3,
    fontFace: "Inter", fontSize: 10, color: terra,
    charSpacing: 14, bold: true,
  });

  s.addText(String(num).padStart(2, "0"), {
    x: cx, y: 0.95, w: cw, h: 0.9,
    fontFace: "Playfair Display", fontSize: 72, color: ochre,
    italic: true, bold: true, valign: "top",
  });

  s.addText(name, {
    x: cx, y: 1.85, w: cw + 0.2, h: 0.7,
    fontFace: "Playfair Display", fontSize: 32, color: ink,
    bold: true,
  });

  s.addText(ruName, {
    x: cx, y: 2.55, w: cw, h: 0.4,
    fontFace: "Playfair Display", fontSize: 18, color: blue,
    italic: true,
  });

  s.addShape(pres.ShapeType.line, {
    x: cx, y: 3.1, w: 1.5, h: 0,
    line: { color: terra, pt: 2 },
  });

  const bodyRuns = [];
  body.forEach((para, i) => {
    if (i > 0) bodyRuns.push({ text: "\n", options: { breakLine: true } });
    bodyRuns.push({ text: para, options: { fontFace: "Inter", fontSize: 12, color: ink } });
  });
  s.addText(bodyRuns, {
    x: cx, y: 3.3, w: cw, h: 2.0,
    valign: "top", paraSpaceAfter: 6,
  });

  // Fact panel
  s.addShape(pres.ShapeType.rect, {
    x: cx, y: 5.4, w: cw, h: 1.2,
    fill: { color: cream }, line: { type: "none" },
  });
  s.addShape(pres.ShapeType.rect, {
    x: cx, y: 5.4, w: 0.08, h: 1.2,
    fill: { color: terra }, line: { type: "none" },
  });

  const factW = (cw - 0.4) / facts.length;
  facts.forEach((fact, j) => {
    const fx = cx + 0.3 + j * factW;
    s.addText(fact.label.toUpperCase(), {
      x: fx, y: 5.55, w: factW, h: 0.3,
      fontFace: "Inter", fontSize: 8, color: muted,
      charSpacing: 8, bold: true,
    });
    s.addText(fact.value, {
      x: fx, y: 5.85, w: factW, h: 0.6,
      fontFace: "Playfair Display", fontSize: 16, color: ink,
      bold: true,
    });
  });

  addPageMark(s, num + 2, total);
}

// ─── SLIDE 1: Cover ───────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: paper };

  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 0.5, h: 7.5,
    fill: { color: terra }, line: { type: "none" },
  });

  s.addShape(pres.ShapeType.ellipse, {
    x: 11.6, y: 0.6, w: 1.2, h: 1.2,
    fill: { color: ochre }, line: { type: "none" },
  });
  s.addText("2026", {
    x: 11.6, y: 0.85, w: 1.2, h: 0.6,
    fontFace: "Playfair Display", fontSize: 24, color: ink,
    bold: true, align: "center",
  });

  s.addText("ПУТЕШЕСТВИЕ  ·  ГИД ПО ГОРОДУ", {
    x: 1.2, y: 1.4, w: 9, h: 0.4,
    fontFace: "Inter", fontSize: 12, color: terra,
    charSpacing: 20, bold: true,
  });

  s.addText("BARCELONA", {
    x: 1.2, y: 1.95, w: 11.5, h: 2.4,
    fontFace: "Playfair Display", fontSize: 110, color: ink,
    charSpacing: -2, bold: true,
  });

  s.addText("Un Día en la Capital Catalana", {
    x: 1.2, y: 4.35, w: 11, h: 0.8,
    fontFace: "Playfair Display", fontSize: 30, color: blue,
    italic: true,
  });

  s.addShape(pres.ShapeType.line, {
    x: 1.2, y: 5.25, w: 4, h: 0,
    line: { color: terra, pt: 2 },
  });

  s.addText([
    { text: "Шесть мест, которые делают город ", options: { color: ink } },
    { text: "незабываемым", options: { color: terra, italic: true, bold: true } },
    { text: ".", options: { color: ink } },
  ], {
    x: 1.2, y: 5.45, w: 10, h: 0.6,
    fontFace: "Inter", fontSize: 18,
  });

  addMosaicChips(s, 9.5, 6.3, 3.2, 0.7, 30, [terra, ochre, blue, cream, muted]);

  s.addText("ГОТОВЫЙ МАРШРУТ  ·  PRACTICAL TIPS  ·  6 ICONIC SPOTS", {
    x: 1.2, y: 7.05, w: 10, h: 0.3,
    fontFace: "Inter", fontSize: 9, color: muted,
    charSpacing: 14,
  });
}

// ─── SLIDE 2: Contents ────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: paper };

  addEyebrow(s, "CONTENIDOS  ·  СОДЕРЖАНИЕ");

  s.addText("Маршрут на день", {
    x: 0.6, y: 0.9, w: 8, h: 1,
    fontFace: "Playfair Display", fontSize: 48, color: ink, bold: true,
  });

  s.addText("Six neighborhoods, one unforgettable day.", {
    x: 0.6, y: 1.85, w: 8, h: 0.5,
    fontFace: "Playfair Display", fontSize: 16, color: blue, italic: true,
  });

  s.addShape(pres.ShapeType.line, {
    x: 0.6, y: 2.5, w: 12.1, h: 0,
    line: { color: terra, pt: 1.5 },
  });

  const items = [
    ["01", "Sagrada Família",  "Шедевр Гауди в строительстве",        "Eixample"],
    ["02", "Park Güell",       "Мозаичный сад на холме",               "Gràcia"],
    ["03", "Casa Batlló",      "Дом-дракон на Passeig de Gràcia",      "Eixample"],
    ["04", "Gothic Quarter",   "Средневековые улочки",                 "Ciutat Vella"],
    ["05", "La Boqueria",      "Главный рынок города",                 "El Raval"],
    ["06", "Barceloneta",      "Пляж и закаты у моря",                 "Barceloneta"],
  ];

  const colX = [0.6, 7.0];
  const rowH = 1.5;
  const startY = 2.85;

  items.forEach(([num, name, ruName, district], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = colX[col];
    const y = startY + row * rowH;

    s.addText(num, {
      x, y, w: 1.1, h: 1.0,
      fontFace: "Playfair Display", fontSize: 46, color: ochre,
      italic: true, bold: true, valign: "top",
    });
    s.addText(name, {
      x: x + 1.2, y: y + 0.05, w: 4.8, h: 0.5,
      fontFace: "Playfair Display", fontSize: 22, color: ink, bold: true,
    });
    s.addText(ruName, {
      x: x + 1.2, y: y + 0.55, w: 4.8, h: 0.35,
      fontFace: "Inter", fontSize: 12, color: blue,
    });
    s.addText(district.toUpperCase(), {
      x: x + 1.2, y: y + 0.9, w: 4.8, h: 0.25,
      fontFace: "Inter", fontSize: 9, color: terra,
      charSpacing: 10, bold: true,
    });
  });

  addPageMark(s, 2, TOTAL);
}

// ─── SLIDES 3–8: Spot slides ──────────────────────────────────────────────────

addSpotSlide({
  num: 1, total: TOTAL,
  name: "Sagrada Família",
  ruName: "Шедевр Гауди в работе с 1882 года",
  district: "Eixample",
  body: [
    "Самый знаменитый собор Барселоны и единственный, который до сих пор строится. Антонио Гауди работал над ним 43 года и похоронен в крипте под зданием.",
    "Внутри — лес из колонн-деревьев и витражи, превращающие стены в калейдоскоп. Свет утром идёт через синие и зелёные стёкла со стороны Рождества, вечером — через красные и золотые со стороны Страстей.",
    "Достроить храм планируют к 2026 году — к 100-летию смерти архитектора.",
  ],
  facts: [
    { label: "Вход",  value: "26 €" },
    { label: "Часы",  value: "9:00–18:00" },
    { label: "Время", value: "1.5 ч" },
  ],
  photoCaption: "Фасад Рождества с башнями, утренний свет",
  photoOnLeft: true,
});

addSpotSlide({
  num: 2, total: TOTAL,
  name: "Park Güell",
  ruName: "Парк-сказка с мозаичной саламандрой",
  district: "Gràcia",
  body: [
    "Изначально Гауди задумывал жилой квартал для богатой буржуазии — затея провалилась, и теперь это парк. Здесь — знаменитая саламандра El Drac, волнистая мозаичная скамья и колоннада Hipóstila.",
    "С верхней террасы открывается лучший панорамный вид на город: от Tibidabo до моря, с Sagrada Família ровно посередине.",
    "В платную «монументальную зону» вход по таймслотам — бронировать заранее обязательно.",
  ],
  facts: [
    { label: "Вход",  value: "18 €" },
    { label: "Метро", value: "Lesseps" },
    { label: "Лучше", value: "На закате" },
  ],
  photoCaption: "Мозаичная скамья и панорама Барселоны",
  photoOnLeft: false,
});

addSpotSlide({
  num: 3, total: TOTAL,
  name: "Casa Batlló",
  ruName: "«Дом костей» на Passeig de Gràcia",
  district: "Eixample",
  body: [
    "Жилой дом, перестроенный Гауди в 1904–1906 годах для семьи текстильного магната Жозепа Батльо. Каталонцы прозвали его Casa dels Ossos — «Дом из костей»: балконы похожи на черепа, колонны — на берцовые кости.",
    "Крыша — это спина дракона, которого пронзает копьё Святого Жоржа, покровителя Каталонии. Внутренний дворик облицован плиткой, темнеющей сверху вниз — Гауди добивался равномерного освещения этажей.",
    "Рядом — Casa Milà («La Pedrera»), ещё один шедевр Гауди в десяти минутах пешком.",
  ],
  facts: [
    { label: "Вход",  value: "от 29 €" },
    { label: "Метро", value: "Passeig de Gràcia" },
    { label: "Время", value: "1 ч" },
  ],
  photoCaption: "Фасад Casa Batlló с балконами-черепами",
  photoOnLeft: true,
});

addSpotSlide({
  num: 4, total: TOTAL,
  name: "Barri Gòtic",
  ruName: "Готический квартал — сердце старого города",
  district: "Ciutat Vella",
  body: [
    "Лабиринт узких улиц между Рамблас и Виа Лаетана. Здесь стоит готический собор Святого Креста (XIII–XV вв.), сохранились римские стены и средневековые еврейские кварталы El Call.",
    "Главные площади — Plaça Reial с пальмами и фонарями работы юного Гауди и Plaça del Rei, где, по легенде, Колумба принимали Фердинанд и Изабелла после возвращения из Америки.",
    "В отличие от современного Эшампле, тут нет правильной сетки — заблудиться приятно и неизбежно.",
  ],
  facts: [
    { label: "Вход",  value: "Бесплатно" },
    { label: "Лучше", value: "Утро" },
    { label: "Стиль", value: "Готика" },
  ],
  photoCaption: "Узкая улочка Готического квартала",
  photoOnLeft: false,
});

addSpotSlide({
  num: 5, total: TOTAL,
  name: "La Boqueria",
  ruName: "Главный рынок Барселоны с 1217 года",
  district: "El Raval",
  body: [
    "Mercat de Sant Josep de la Boqueria — рынок с восьмисотлетней историей, открытый прямо на Рамблас. Под чугунным навесом 1840 года — больше 200 прилавков с хамоном, морепродуктами, фруктами и тапасами.",
    "Лучше приходить голодным к 11:00 — обязательные пробы: pa amb tomàquet (хлеб с томатом), iberico (хамон), и стакан кавы (каталонское игристое).",
    "Лучшие bar-стойки — у дальнего входа, не у фасада: там меньше туристов и свежее цены.",
  ],
  facts: [
    { label: "Вход",    value: "Бесплатно" },
    { label: "Часы",    value: "8:00–20:30" },
    { label: "Закрыт",  value: "Воскресенье" },
  ],
  photoCaption: "Прилавок с хамоном и оливками",
  photoOnLeft: true,
});

addSpotSlide({
  num: 6, total: TOTAL,
  name: "Barceloneta",
  ruName: "Старый рыбацкий квартал у моря",
  district: "Barceloneta",
  body: [
    "Бывшая рыбацкая слобода XVIII века, разлинованная узкими улочками-параллелями. Сейчас — самый живой пляжный район города с парадом баров, паэлий и сёрфингистов.",
    "К пляжу выходит набережная Passeig Marítim длиной 4 км — идеально для вечерней пробежки или велопроката. На горизонте — гигантский золотой Peix Фрэнка Гери, «Рыба» с Олимпийских игр 1992 года.",
    "Закатные тапас в chiringuito на песке — лучший финал дня.",
  ],
  facts: [
    { label: "Вход",  value: "Бесплатно" },
    { label: "Метро", value: "Barceloneta" },
    { label: "Лучше", value: "Закат" },
  ],
  photoCaption: "Пляж Barceloneta на закате",
  photoOnLeft: false,
});

// ─── SLIDE 9: Tips ────────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: paper };

  addEyebrow(s, "PRÁCTICA  ·  ПРАКТИЧЕСКАЯ ИНФОРМАЦИЯ");

  s.addText("Полезно знать", {
    x: 0.6, y: 0.9, w: 8, h: 1,
    fontFace: "Playfair Display", fontSize: 44, color: ink, bold: true,
  });

  s.addText("Если едете впервые.", {
    x: 0.6, y: 1.85, w: 8, h: 0.5,
    fontFace: "Playfair Display", fontSize: 16, color: blue, italic: true,
  });

  s.addShape(pres.ShapeType.line, {
    x: 0.6, y: 2.5, w: 12.1, h: 0,
    line: { color: terra, pt: 1.5 },
  });

  const cards = [
    {
      icon: "→",
      title: "Когда ехать",
      text: "Лучшие месяцы — апрель–июнь и сентябрь–октябрь. Июль–август жарко и переполнено; зимой мягко (12–15 °C), но море уже холодное.",
    },
    {
      icon: "€",
      title: "Деньги и оплата",
      text: "Евро. Безналичные платежи везде, кроме мелких баров. Турналог €4 / ночь в отеле. Чаевые — 5–10 % если без service charge.",
    },
    {
      icon: "M",
      title: "Транспорт",
      text: "Метро 11 линий, T-casual на 10 поездок — €12.55. Из аэропорта El Prat в центр — Aerobús €7,25 или метро L9 Sud €5,70.",
    },
    {
      icon: "!",
      title: "Безопасность",
      text: "Карманники в метро, на Рамблас и у Sagrada Família. Рюкзак держите спереди в толпе. Скорая — 112, туристическая полиция — у Plaça Catalunya.",
    },
  ];

  const cardW = 2.9;
  const cardH = 3.7;
  const cardY = 2.85;
  const gap = 0.15;
  const totalW = cardW * 4 + gap * 3;
  const startX = (SLIDE_W - totalW) / 2;
  const accentColors = [terra, ochre, blue, terra];

  cards.forEach((card, i) => {
    const x = startX + i * (cardW + gap);
    const ac = accentColors[i];

    s.addShape(pres.ShapeType.rect, {
      x, y: cardY, w: cardW, h: cardH,
      fill: { color: cream }, line: { type: "none" },
      shadow: shadow(0.1),
    });
    s.addShape(pres.ShapeType.rect, {
      x, y: cardY, w: cardW, h: 0.12,
      fill: { color: ac }, line: { type: "none" },
    });
    s.addText(card.icon, {
      x: x + 0.3, y: cardY + 0.4, w: 1, h: 1,
      fontFace: "Playfair Display", fontSize: 60, color: ac,
      italic: true, bold: true, align: "left",
    });
    s.addText(card.title, {
      x: x + 0.3, y: cardY + 1.5, w: cardW - 0.6, h: 0.7,
      fontFace: "Playfair Display", fontSize: 18, color: ink,
      bold: true, valign: "top",
    });
    s.addShape(pres.ShapeType.line, {
      x: x + 0.3, y: cardY + 2.15, w: 0.6, h: 0,
      line: { color: terra, pt: 1.5 },
    });
    s.addText(card.text, {
      x: x + 0.3, y: cardY + 2.3, w: cardW - 0.6, h: 1.3,
      fontFace: "Inter", fontSize: 10.5, color: ink,
      valign: "top",
    });
  });

  s.addText("¡Bon viatge!", {
    x: 0.6, y: 6.65, w: 4, h: 0.4,
    fontFace: "Playfair Display", fontSize: 22, color: terra,
    italic: true, bold: true,
  });

  addMosaicChips(s, 5.5, 6.65, 6.4, 0.4, 22, [terra, ochre, blue, muted]);

  addPageMark(s, 9, TOTAL);
}

// ─── Output ───────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "Barcelona_Travel.pptx" })
  .then(() => console.log("Done: Barcelona_Travel.pptx"))
  .catch(err => { console.error(err); process.exit(1); });
