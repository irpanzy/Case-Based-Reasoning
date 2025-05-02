const xlsx = require("xlsx"); //Import library xlsx untuk membaca dan menulis file Excel

function bacaData(filename) {
  const workbook = xlsx.readFile(filename); // Membaca file Excel
  const sheetName = workbook.SheetNames[0]; // Mengambil nama sheet pertama
  const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]); // Konversi ke array ke JSON

  console.log(data); // Menampilkan data untuk debugging

  return data;
}

function fuzzyServis(pelayanan) {
  return {
    buruk: pelayanan <= 50 ? 1 - pelayanan / 50 : 0,
    sedang:
      pelayanan > 50 && pelayanan < 70
        ? pelayanan <= 60
          ? (pelayanan - 50) / 10
          : (70 - pelayanan) / 10
        : 0,
    baik: pelayanan >= 70 ? (pelayanan - 70) / 10 : 0,
  };
}

function fuzzyHarga(harga) {
  return {
    murah: harga <= 40000 ? 1 - (harga - 25000) / 15000 : 0,
    sedang:
      harga > 40000 && harga < 50000
        ? harga <= 45000
          ? (harga - 40000) / 5000
          : (50000 - harga) / 5000
        : 0,
    mahal: harga >= 50000 ? (harga - 50000) / 2500 : 0,
  };
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

function inferensi(servisFuzzy, hargaFuzzy) {
  const rules = [];

  // 3 (servis) × 3 (harga) = 9 aturan
  rules.push({ skor: Math.min(servisFuzzy.baik, hargaFuzzy.murah), nilai: 90 });
  rules.push({ skor: Math.min(servisFuzzy.baik, hargaFuzzy.sedang), nilai: 80 });
  rules.push({ skor: Math.min(servisFuzzy.baik, hargaFuzzy.mahal), nilai: 70 });

  rules.push({ skor: Math.min(servisFuzzy.sedang, hargaFuzzy.murah), nilai: 70 });
  rules.push({ skor: Math.min(servisFuzzy.sedang, hargaFuzzy.sedang), nilai: 60 });
  rules.push({ skor: Math.min(servisFuzzy.sedang, hargaFuzzy.mahal), nilai: 50 });

  rules.push({ skor: Math.min(servisFuzzy.buruk, hargaFuzzy.murah), nilai: 50 });
  rules.push({ skor: Math.min(servisFuzzy.buruk, hargaFuzzy.sedang), nilai: 40 });
  rules.push({ skor: Math.min(servisFuzzy.buruk, hargaFuzzy.mahal), nilai: 30 });

  return rules;
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
