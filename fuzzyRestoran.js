const xlsx = require("xlsx");

function bacaData(filename) {
  const workbook = xlsx.readFile(filename);
  const sheetName = workbook.SheetNames[0];
  const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  console.log(data);

  return data;
}

function fuzzyServis(pelayanan) {
  return {
    buruk: pelayanan <= 50 ? 1 - pelayanan / 50 : 0,
    sedang:
      pelayanan > 30 && pelayanan < 70
        ? pelayanan <= 50
          ? (pelayanan - 30) / 20
          : (70 - pelayanan) / 20
        : 0,
    baik: pelayanan >= 50 ? (pelayanan - 50) / 50 : 0,
  };
}

function fuzzyHarga(harga) {
  return {
    murah: harga <= 40000 ? 1 - (harga - 25000) / 15000 : 0,
    sedang:
      harga > 30000 && harga < 50000
        ? harga <= 40000
          ? (harga - 30000) / 10000
          : (50000 - harga) / 10000
        : 0,
    mahal: harga >= 40000 ? (harga - 40000) / 15000 : 0,
  };
}

function inferensi(servisFuzzy, hargaFuzzy) {
  const rules = [];

  rules.push({
    skor: Math.min(servisFuzzy.baik, hargaFuzzy.murah),
    nilai: 90,
  });

  rules.push({
    skor: Math.min(servisFuzzy.baik, hargaFuzzy.sedang),
    nilai: 80,
  });

  rules.push({
    skor: Math.min(servisFuzzy.sedang, hargaFuzzy.murah),
    nilai: 70,
  });

  rules.push({
    skor: Math.min(servisFuzzy.sedang, hargaFuzzy.sedang),
    nilai: 60,
  });

  rules.push({
    skor: Math.min(servisFuzzy.buruk, hargaFuzzy.murah),
    nilai: 50,
  });

  return rules;
}

function defuzzification(rules) {
  let atas = 0;
  let bawah = 0;

  rules.forEach((rule) => {
    atas += rule.skor * rule.nilai;
    bawah += rule.skor;
  });

  if (bawah === 0) return 0;
  return atas / bawah;
}

function prosesFuzzy(inputFile, outputFile) {
  const dataRestoran = bacaData(inputFile);

  const hasil = dataRestoran.map((resto, idx) => {
    const pelayanan = resto.Pelayanan;
    const harga = resto.harga;

    const servisFuzzy = fuzzyServis(pelayanan);
    const hargaFuzzy = fuzzyHarga(harga);

    const rules = inferensi(servisFuzzy, hargaFuzzy);
    const skor = defuzzification(rules);

    return {
      ID_Restoran: resto["id Pelanggan"],
      Kualitas_Servis: pelayanan,
      Harga: harga,
      Skor: skor,
    };
  });

  hasil.sort((a, b) => b.Skor - a.Skor);

  const top5 = hasil.slice(0, 5);

  const ws = xlsx.utils.json_to_sheet(top5);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Top5Restoran");
  xlsx.writeFile(wb, outputFile);

  console.log("Proses selesai! File output:", outputFile);

  console.log("Daftar 5 Restoran Terbaik:");
  console.table(top5);
}

prosesFuzzy("restoran.xlsx", "peringkat.xlsx");
