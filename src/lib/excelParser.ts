import * as XLSX from 'xlsx';

export interface RespondentRecord {
  id: string;
  q: number[]; // 35 skor (skala 1-5)
}

export interface HabitCategoryStat {
  index: number;
  name: string;
  shortName: string;
  range: string;
  mean: number;
  totalMean: number;
  median: number;
  stdDev: number;
  predikat: string;
  distPct: [number, number, number, number, number]; // % skala 1..5
  itemMeans: { qCode: string; label: string; mean: number }[];
  topBarrier: { qCode: string; label: string; mean: number };
  topStrength: { qCode: string; label: string; mean: number };
}

export interface AnomalyInsight {
  type: 'warning' | 'paradox' | 'bias' | 'info';
  title: string;
  metric: string;
  description: string;
  respondents: string[];
}

export interface DashboardAnalytics {
  totalRespondents: number;
  overallMean: number;
  overallTotalMean: number;
  overallPredikat: string;
  predikatCounts: Record<string, number>;
  habits: HabitCategoryStat[];
  sortedHabits: HabitCategoryStat[];
  correlationMatrix: number[][];
  anomalies: AnomalyInsight[];
  respondentProfiles: {
    id: string;
    habitMeans: number[];
    overallMean: number;
    overallPredikat: string;
  }[];
}

export const HABIT_DEFINITIONS = [
  { name: 'Bangun Pagi', shortName: 'Bangun Pagi', start: 0 },
  { name: 'Beribadah', shortName: 'Beribadah', start: 5 },
  { name: 'Berolahraga', shortName: 'Berolahraga', start: 10 },
  { name: 'Makan Sehat dan Bergizi', shortName: 'Makan Sehat', start: 15 },
  { name: 'Gemar Belajar', shortName: 'Gemar Belajar', start: 20 },
  { name: 'Bermasyarakat', shortName: 'Bermasyarakat', start: 25 },
  { name: 'Tidur Cepat', shortName: 'Tidur Cepat', start: 30 },
];

export const QUESTION_LABELS: string[] = [
  'Bangun pagi mandiri tanpa dibangunkan berkali-kali',
  'Merapikan tempat tidur & bersiap lebih awal',
  'Aktivitas ringan setelah bangun pagi agar segar',
  'Datang ke sekolah lebih awal agar tenang',
  'Bangun pagi tepat waktu untuk ibadah pagi',
  'Melaksanakan ibadah wajib secara disiplin',
  'Mengawali dan mengakhiri kegiatan dengan berdoa',
  'Menghormati teman yang beribadah agama lain',
  'Antusias mengikuti hari besar keagamaan',
  'Ajaran agama mendorong sikap jujur dan santun',
  'Rutin berolahraga minimal 2-3 kali seminggu',
  'Aktif dalam ekstrakurikuler/lomba olahraga',
  'Peregangan/aktivitas fisik ringan saat istirahat',
  'Olahraga melatih disiplin, suportif & pantang menyerah',
  'Menggunakan & merawat sarana olahraga sekolah',
  'Rutin mengonsumsi sayur dan buah setiap hari',
  'Selalu sarapan pagi sebelum berangkat sekolah',
  'Membawa bekal dengan wadah guna ulang',
  'Membatasi jajanan kemasan berpengawet/pemanis',
  'Disiplin makan teratur tepat waktu setiap hari',
  'Membaca buku literasi setiap hari',
  'Memanfaatkan internet positif untuk sumber belajar',
  'Rajin mengulas materi pelajaran mandiri',
  'Proaktif mencari solusi saat kesulitan tugas/P5',
  'Mengerjakan ujian & tugas secara jujur',
  'Partisipasi aktif kerja kelompok & gotong royong',
  'Membuang sampah pada tempatnya',
  'Berteman tanpa membedakan latar belakang',
  'Ikut merawat tanaman penghijauan sekolah',
  'Terbiasa mengucapkan tolong, maaf, terima kasih',
  'Jadwal tidur malam teratur & hindari begadang',
  'Berhenti menggunakan gadget 30 menit sebelum tidur',
  'Memastikan waktu tidur cukup untuk kesehatan',
  'Konsentrasi & daya ingat baik berkat tidur cukup',
  'Istirahat cukup membantu kendalikan emosi',
];

export function getPredikat(score: number): string {
  if (score <= 1.8) return 'Sangat Buruk';
  if (score <= 2.6) return 'Buruk';
  if (score <= 3.4) return 'Cukup';
  if (score <= 4.2) return 'Baik';
  return 'Sangat Baik';
}

function pearsonCorrelation(x: number[], y: number[]): number {
  const n = x.length;
  if (n === 0) return 0;
  const meanX = x.reduce((a, b) => a + b, 0) / n;
  const meanY = y.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let denX = 0;
  let denY = 0;
  for (let i = 0; i < n; i++) {
    const dx = x[i] - meanX;
    const dy = y[i] - meanY;
    num += dx * dy;
    denX += dx * dx;
    denY += dy * dy;
  }
  const denom = Math.sqrt(denX * denY);
  return denom === 0 ? 0 : Number((num / denom).toFixed(3));
}

export async function parseExcelFile(file: File): Promise<RespondentRecord[]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  // Prioritaskan sheet 'Data Bersih & Skor Individu' jika ada, jika tidak ambil sheet pertama
  const targetSheetName =
    workbook.SheetNames.find((s) => s.toLowerCase().includes('data bersih')) ||
    workbook.SheetNames[0];
  const sheet = workbook.Sheets[targetSheetName];
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // Cari baris awal data numerik (lewati header banner)
  const records: RespondentRecord[] = [];
  let counter = 1;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length < 35) continue;

    // Ekstrak seluruh nilai valid skala 1..5 pada baris tersebut
    const numericScores: number[] = [];
    for (let c = 0; c < row.length; c++) {
      const val = Number(row[c]);
      if (Number.isInteger(val) && val >= 1 && val <= 5) {
        numericScores.push(val);
      }
      if (numericScores.length === 35) break;
    }

    if (numericScores.length === 35) {
      const rawId =
        typeof row[0] === 'string' && row[0].toLowerCase().includes('responden')
          ? row[0].trim()
          : `Responden ${counter}`;
      records.push({ id: rawId, q: numericScores });
      counter++;
    }
  }

  if (records.length === 0) {
    throw new Error(
      'Format Excel tidak dikenali. Pastikan file memiliki 35 kolom skor pertanyaan (skala 1–5).'
    );
  }
  return records;
}

export function computeAnalytics(records: RespondentRecord[]): DashboardAnalytics {
  const n = records.length;

  const respondentProfiles = records.map((rec) => {
    const habitMeans = HABIT_DEFINITIONS.map((h) => {
      const slice = rec.q.slice(h.start, h.start + 5);
      return Number((slice.reduce((a, b) => a + b, 0) / 5).toFixed(2));
    });
    const overallMean = Number((rec.q.reduce((a, b) => a + b, 0) / 35).toFixed(2));
    return {
      id: rec.id,
      habitMeans,
      overallMean,
      overallPredikat: getPredikat(overallMean),
    };
  });

  const predikatCounts: Record<string, number> = {
    'Sangat Baik': 0,
    Baik: 0,
    Cukup: 0,
    Buruk: 0,
    'Sangat Buruk': 0,
  };
  respondentProfiles.forEach((p) => {
    predikatCounts[p.overallPredikat] = (predikatCounts[p.overallPredikat] || 0) + 1;
  });

  const habits: HabitCategoryStat[] = HABIT_DEFINITIONS.map((h, hIdx) => {
    const catScores = respondentProfiles.map((p) => p.habitMeans[hIdx]);
    const sorted = [...catScores].sort((a, b) => a - b);
    const mean = Number((catScores.reduce((a, b) => a + b, 0) / n).toFixed(2));
    const median =
      n % 2 === 0
        ? Number(((sorted[n / 2 - 1] + sorted[n / 2]) / 2).toFixed(2))
        : sorted[Math.floor(n / 2)];
    const variance =
      catScores.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / Math.max(n - 1, 1);
    const stdDev = Number(Math.sqrt(variance).toFixed(2));

    // Distribusi skala 1..5 pada 5 butir pertanyaan
    const counts = [0, 0, 0, 0, 0];
    for (const rec of records) {
      for (let qIdx = h.start; qIdx < h.start + 5; qIdx++) {
        const v = rec.q[qIdx];
        if (v >= 1 && v <= 5) counts[v - 1]++;
      }
    }
    const totalItems = n * 5;
    const distPct = counts.map((c) => Number(((c / totalItems) * 100).toFixed(1))) as [
      number,
      number,
      number,
      number,
      number,
    ];

    const itemMeans = Array.from({ length: 5 }, (_, i) => {
      const qIdx = h.start + i;
      const sum = records.reduce((acc, r) => acc + r.q[qIdx], 0);
      return {
        qCode: `Q${qIdx + 1}`,
        label: QUESTION_LABELS[qIdx],
        mean: Number((sum / n).toFixed(2)),
      };
    });

    const sortedItems = [...itemMeans].sort((a, b) => a.mean - b.mean);

    return {
      index: hIdx + 1,
      name: h.name,
      shortName: h.shortName,
      range: `Q${h.start + 1}–Q${h.start + 5}`,
      mean,
      totalMean: Number((mean * 5).toFixed(2)),
      median,
      stdDev,
      predikat: getPredikat(mean),
      distPct,
      itemMeans,
      topBarrier: sortedItems[0],
      topStrength: sortedItems[sortedItems.length - 1],
    };
  });

  const sortedHabits = [...habits].sort((a, b) => b.mean - a.mean);
  const overallMean = Number(
    (habits.reduce((acc, h) => acc + h.mean, 0) / habits.length).toFixed(2)
  );
  const overallTotalMean = Number((overallMean * 35).toFixed(2));

  // Matriks Korelasi Antar-Kebiasaan (7x7)
  const correlationMatrix = habits.map((_, i) =>
    habits.map((__, j) =>
      pearsonCorrelation(
        respondentProfiles.map((p) => p.habitMeans[i]),
        respondentProfiles.map((p) => p.habitMeans[j])
      )
    )
  );

  // Deteksi Anomali & Insight Otomatis
  const anomalies: AnomalyInsight[] = [];

  const straight5 = records.filter((r) => r.q.every((v) => v === 5)).map((r) => r.id);
  if (straight5.length > 0) {
    anomalies.push({
      type: 'bias',
      title: 'Pola Straight-Lining (Skor Sempurna 5,00 Tanpa Variasi)',
      metric: `${straight5.length} Siswa`,
      description:
        'Responden mengisi nilai maksimal (5) pada seluruh 35 indikator. Indikasi potensi social desirability bias yang perlu diverifikasi melalui observasi guru.',
      respondents: straight5,
    });
  }

  const lowRisk = respondentProfiles
    .filter((p) => p.overallMean <= 1.8)
    .map((p) => `${p.id} (${p.overallMean})`);
  if (lowRisk.length > 0) {
    anomalies.push({
      type: 'warning',
      title: 'Kelompok Risiko Ekstrem Rendah (Predikat Sangat Buruk)',
      metric: `${lowRisk.length} Siswa`,
      description:
        'Skor rata-rata keseluruhan di bawah 1,80 pada hampir seluruh dimensi kebiasaan. Memerlukan pendampingan konseling personal oleh Wali Kelas/BK.',
      respondents: lowRisk,
    });
  }

  const sleepParadox = respondentProfiles
    .filter((p) => p.habitMeans[1] >= 4.6 && p.habitMeans[6] <= 2.4)
    .map((p) => `${p.id} (Ibadah ${p.habitMeans[1]} vs Tidur ${p.habitMeans[6]})`);
  if (sleepParadox.length > 0) {
    anomalies.push({
      type: 'paradox',
      title: 'Paradoks "Taat Ibadah tetapi Kurang Tidur"',
      metric: `${sleepParadox.length} Siswa`,
      description:
        'Memiliki kedisiplinan spiritual sangat tinggi (>= 4,60) namun skor Tidur Cepat rendah (<= 2,40) akibat penggunaan gadget larut malam.',
      respondents: sleepParadox,
    });
  }

  const sedentaryAchievers = respondentProfiles
    .filter((p) => p.habitMeans[4] >= 4.4 && p.habitMeans[2] <= 2.2)
    .map((p) => `${p.id} (Belajar ${p.habitMeans[4]} vs Olahraga ${p.habitMeans[2]})`);
  if (sedentaryAchievers.length > 0) {
    anomalies.push({
      type: 'info',
      title: 'Profil Sedentary Achiever (Akademik Tinggi, Minim Gerak)',
      metric: `${sedentaryAchievers.length} Siswa`,
      description:
        'Sangat tekun belajar mandiri (>= 4,40) tetapi sangat jarang berolahraga (<= 2,20). Berisiko mengalami kelelahan fisik dan postur kaku.',
      respondents: sedentaryAchievers,
    });
  }

  return {
    totalRespondents: n,
    overallMean,
    overallTotalMean,
    overallPredikat: getPredikat(overallMean),
    predikatCounts,
    habits,
    sortedHabits,
    correlationMatrix,
    anomalies,
    respondentProfiles,
  };
    }
           
