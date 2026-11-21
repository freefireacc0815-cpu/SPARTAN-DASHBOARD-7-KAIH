'use client';

import React, { useState, useMemo } from 'react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  Radar,
  Cell,
  Legend,
} from 'recharts';
import {
  Upload,
  Users,
  Award,
  TrendingUp,
  AlertTriangle,
  CheckCircle2,
  Sparkles,
  Search,
  BarChart3,
  Layers,
  Activity,
  ShieldAlert,
  FileSpreadsheet,
} from 'lucide-react';
import {
  computeAnalytics,
  parseExcelFile,
  RespondentRecord,
  HABIT_DEFINITIONS,
} from '../lib/excelParser';
import { DEFAULT_SMAN1_DATA } from '../lib/defaultData';

export default function SpartanDashboard() {
  const [records, setRecords] = useState<RespondentRecord[]>(DEFAULT_SMAN1_DATA);
  const [fileName, setFileName] = useState<string>('Sampel SMAN_1_PANDEGLANG.xlsx (Unggah file penuh untuk 233 siswa)');
  const [activeTab, setActiveTab] = useState<'overview' | 'habits' | 'patterns' | 'anomalies'>('overview');
  const [selectedHabitIdx, setSelectedHabitIdx] = useState<number>(6); // Default sorot Berolahraga / terendah
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [filterPredikat, setFilterPredikat] = useState<string>('ALL');
  const [uploadError, setUploadError] = useState<string | null>(null);

  const analytics = useMemo(() => computeAnalytics(records), [records]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      setUploadError(null);
      const parsed = await parseExcelFile(file);
      setRecords(parsed);
      setFileName(`${file.name} (${parsed.length} Responden)`);
    } catch (err: any) {
      setUploadError(err.message || 'Gagal memproses file Excel.');
    }
  };

  const filteredProfiles = useMemo(() => {
    return analytics.respondentProfiles.filter((p) => {
      const matchId = p.id.toLowerCase().includes(searchTerm.toLowerCase());
      const matchPred = filterPredikat === 'ALL' || p.overallPredikat === filterPredikat;
      return matchId && matchPred;
    });
  }, [analytics.respondentProfiles, searchTerm, filterPredikat]);

  const topHabit = analytics.sortedHabits[0];
  const lowestHabit = analytics.sortedHabits[analytics.sortedHabits.length - 1];
  const currentHabit = analytics.habits[selectedHabitIdx] || analytics.habits[0];

  const predikatBadgeColor = (pred: string) => {
    switch (pred) {
      case 'Sangat Baik':
        return 'bg-emerald-100 text-emerald-800 border-emerald-200';
      case 'Baik':
        return 'bg-blue-100 text-blue-800 border-blue-200';
      case 'Cukup':
        return 'bg-amber-100 text-amber-800 border-amber-200';
      default:
        return 'bg-rose-100 text-rose-800 border-rose-200';
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      {/* Top Navbar */}
      <header className="bg-gradient-to-r from-blue-900 via-blue-800 to-emerald-800 text-white shadow-md sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-orange-500 text-white font-extrabold px-3 py-2 rounded-xl shadow-inner tracking-wider text-lg">
              SPARTAN
            </div>
            <div>
              <h1 className="text-xl sm:text-2xl font-bold tracking-tight">
                7 KAIH — IN REAL LIFE ANALYTICS
              </h1>
              <p className="text-xs sm:text-sm text-blue-100">
                Dashboard Evaluasi Implementasi 7 Kebiasaan Anak Indonesia Hebat
              </p>
            </div>
          </div>

          {/* Excel Upload Button */}
          <div className="flex flex-col sm:flex-row items-start sm:items-center gap-2">
            <div className="text-xs bg-white/10 backdrop-blur-sm px-3 py-1.5 rounded-lg border border-white/20 flex items-center gap-2">
              <FileSpreadsheet className="w-4 h-4 text-emerald-300" />
              <span className="truncate max-w-[240px]">{fileName}</span>
            </div>
            <label className="cursor-pointer bg-orange-500 hover:bg-orange-600 transition text-white text-sm font-semibold px-4 py-2 rounded-xl shadow flex items-center gap-2">
              <Upload className="w-4 h-4" />
              <span>Unggah Excel (.xlsx)</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        {uploadError && (
          <div className="bg-rose-50 border border-rose-200 text-rose-800 px-4 py-3 rounded-xl flex items-center gap-2 text-sm">
            <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0" />
            <span>{uploadError}</span>
          </div>
        )}

        {/* Executive KPI Cards */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-200/80 flex items-start justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                Total Partisipan
              </p>
              <p className="text-3xl font-extrabold text-slate-900 mt-1">
                {analytics.totalRespondents}
              </p>
              <p className="text-xs text-slate-500 mt-1">Siswa Teranonimisasi (100% Valid)</p>
            </div>
            <div className="p-3 bg-blue-50 text-blue-700 rounded-xl">
              <Users className="w-6 h-6" />
            </div>
          </div>

          <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-200/80 flex items-start justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                Indeks Kepatuhan 7 KAIH
              </p>
              <div className="flex items-baseline gap-2 mt-1">
                <p className="text-3xl font-extrabold text-blue-700">
                  {analytics.overallMean.toFixed(2)}
                </p>
                <span className="text-sm text-slate-400 font-medium">/ 5,00</span>
              </div>
              <span
                className={`inline-block mt-1.5 text-xs font-semibold px-2.5 py-0.5 rounded-full border ${predikatBadgeColor(
                  analytics.overallPredikat
                )}`}
              >
                Predikat: {analytics.overallPredikat}
              </span>
            </div>
            <div className="p-3 bg-emerald-50 text-emerald-700 rounded-xl">
              <Award className="w-6 h-6" />
            </div>
          </div>

          <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-200/80 flex items-start justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                Kebiasaan Paling Dominan
              </p>
              <p className="text-lg font-bold text-emerald-700 mt-1">{topHabit.name}</p>
              <p className="text-xs text-slate-600 mt-1">
                Skor Rata-rata: <strong className="text-slate-900">{topHabit.mean}</strong> ({topHabit.predikat})
              </p>
            </div>
            <div className="p-3 bg-emerald-50 text-emerald-600 rounded-xl">
              <TrendingUp className="w-6 h-6" />
            </div>
          </div>

          <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-200/80 flex items-start justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">
                Prioritas Intervensi
              </p>
              <p className="text-lg font-bold text-orange-600 mt-1">{lowestHabit.name}</p>
              <p className="text-xs text-slate-600 mt-1">
                Skor Rata-rata: <strong className="text-slate-900">{lowestHabit.mean}</strong> (Gap -
                {(topHabit.mean - lowestHabit.mean).toFixed(2)})
              </p>
            </div>
            <div className="p-3 bg-orange-50 text-orange-600 rounded-xl">
              <AlertTriangle className="w-6 h-6" />
            </div>
          </div>
        </section>

        {/* Navigation Tabs */}
        <nav className="flex flex-wrap gap-2 bg-white p-1.5 rounded-2xl shadow-sm border border-slate-200/80">
          {[
            { id: 'overview', label: '1. Ringkasan Eksekutif', icon: BarChart3 },
            { id: 'habits', label: '2. Analisis Per Kebiasaan (7 KAIH)', icon: Layers },
            { id: 'patterns', label: '3. Pola Tren & Korelasi', icon: Activity },
            { id: 'anomalies', label: '4. Temuan Unik & Profil Siswa', icon: Sparkles },
          ].map((tab) => {
            const Icon = tab.icon;
            const active = activeTab === tab.id;
            return (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id as any)}
                className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-sm font-semibold transition ${
                  active
                    ? 'bg-blue-700 text-white shadow-sm'
                    : 'text-slate-600 hover:bg-slate-100'
                }`}
              >
                <Icon className="w-4 h-4" />
                <span>{tab.label}</span>
              </button>
            );
          })}
        </nav>

        {/* TAB 1: EXECUTIVE OVERVIEW */}
        {activeTab === 'overview' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            {/* Bar Chart Peringkat 7 Kebiasaan */}
            <div className="lg:col-span-7 bg-white p-6 rounded-2xl shadow-sm border border-slate-200/80">
              <h2 className="text-base font-bold text-slate-900">
                Peringkat Skor Rata-rata 7 Kebiasaan (Skala 1–5)
              </h2>
              <p className="text-xs text-slate-500 mb-4">
                Hijau: Sangat Baik (&gt;4,20) | Biru: Baik (3,41–4,20) | Oranye: Perlu Perhatian
              </p>
              <div className="h-80">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={analytics.sortedHabits}
                    layout="vertical"
                    margin={{ top: 5, right: 30, left: 30, bottom: 5 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                    <XAxis type="number" domain={[0, 5]} tick={{ fontSize: 12 }} />
                    <YAxis
                      type="category"
                      dataKey="shortName"
                      width={120}
                      tick={{ fontSize: 12, fontWeight: 600 }}
                    />
                    <Tooltip
                      formatter={(val: any) => [`${val} / 5,00`, 'Rata-rata Skor']}
                      contentStyle={{ borderRadius: '12px', borderColor: '#CBD5E1' }}
                    />
                    <Bar dataKey="mean" radius={[0, 8, 8, 0]} barSize={26}>
                      {analytics.sortedHabits.map((entry, idx) => (
                        <Cell
                          key={`cell-${idx}`}
                          fill={
                            entry.mean > 4.2
                              ? '#059669'
                              : entry.mean >= 3.75
                              ? '#1D4ED8'
                              : '#F97316'
                          }
                        />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Radar Chart Keseimbangan 7 KAIH */}
            <div className="lg:col-span-5 bg-white p-6 rounded-2xl shadow-sm border border-slate-200/80">
              <h2 className="text-base font-bold text-slate-900">
                Peta Keseimbangan Karakter (Radar 7 KAIH)
              </h2>
              <p className="text-xs text-slate-500 mb-4">
                Visualisasi simetri penerapan dimensi spiritual, sosial, akademik, dan fisik
              </p>
              <div className="h-80">
                <ResponsiveContainer width="100%" height="100%">
                  <RadarChart cx="50%" cy="50%" outerRadius="72%" data={analytics.habits}>
                    <PolarGrid stroke="#CBD5E1" />
                    <PolarAngleAxis dataKey="shortName" tick={{ fontSize: 11, fontWeight: 600 }} />
                    <PolarRadiusAxis angle={30} domain={[0, 5]} tick={{ fontSize: 10 }} />
                    <Radar
                      name="Rata-rata Skor"
                      dataKey="mean"
                      stroke="#1D4ED8"
                      fill="#1D4ED8"
                      fillOpacity={0.28}
                    />
                    <Tooltip />
                  </RadarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: DEEP DIVE PER KEBIASAAN */}
        {activeTab === 'habits' && (
          <div className="space-y-6">
            {/* Pill Selector */}
            <div className="flex flex-wrap gap-2">
              {analytics.habits.map((h, idx) => (
                <button
                  key={h.name}
                  onClick={() => setSelectedHabitIdx(idx)}
                  className={`px-4 py-2 rounded-xl text-xs sm:text-sm font-semibold border transition ${
                    selectedHabitIdx === idx
                      ? 'bg-emerald-700 text-white border-emerald-700 shadow-sm'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  {idx + 1}. {h.shortName} ({h.mean})
                </button>
              ))}
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              {/* Detail Indikator */}
              <div className="lg:col-span-8 bg-white p-6 rounded-2xl shadow-sm border border-slate-200/80">
                <div className="flex flex-wrap items-center justify-between gap-2 mb-4">
                  <div>
                    <h2 className="text-lg font-bold text-slate-900">
                      Bedah Indikator: {currentHabit.name} ({currentHabit.range})
                    </h2>
                    <p className="text-xs text-slate-500">
                      Rata-rata Kategori: <strong>{currentHabit.mean}</strong> | Median:{' '}
                      <strong>{currentHabit.median}</strong> | Std. Deviasi:{' '}
                      <strong>{currentHabit.stdDev}</strong>
                    </p>
                  </div>
                  <span
                    className={`text-xs font-bold px-3 py-1 rounded-full border ${predikatBadgeColor(
                      currentHabit.predikat
                    )}`}
                  >
                    {currentHabit.predikat}
                  </span>
                </div>

                <div className="space-y-3">
                  {currentHabit.itemMeans.map((item) => {
                    const pct = (item.mean / 5) * 100;
                    const isLowest = item.qCode === currentHabit.topBarrier.qCode;
                    return (
                      <div
                        key={item.qCode}
                        className={`p-3.5 rounded-xl border ${
                          isLowest ? 'bg-orange-50/70 border-orange-200' : 'bg-slate-50 border-slate-200/70'
                        }`}
                      >
                        <div className="flex items-center justify-between text-xs sm:text-sm font-semibold mb-1.5">
                          <span className="text-slate-800">
                            <span className="text-blue-700 font-bold mr-1.5">{item.qCode}:</span>
                            {item.label}
                          </span>
                          <span className="font-extrabold text-slate-900 ml-3 shrink-0">
                            {item.mean.toFixed(2)}
                          </span>
                        </div>
                        <div className="w-full bg-slate-200 rounded-full h-2.5 overflow-hidden">
                          <div
                            className={`h-2.5 rounded-full ${
                              isLowest ? 'bg-orange-500' : 'bg-blue-600'
                            }`}
                            style={{ width: `${pct}%` }}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Hambatan Terbesar & Kekuatan Utama */}
              <div className="lg:col-span-4 space-y-4">
                <div className="bg-orange-50 border border-orange-200 p-5 rounded-2xl">
                  <div className="flex items-center gap-2 text-orange-800 font-bold text-sm mb-2">
                    <AlertTriangle className="w-5 h-5 text-orange-600" />
                    <span>Hambatan Terbesar ({currentHabit.topBarrier.qCode})</span>
                  </div>
                  <p className="text-sm font-semibold text-slate-900 mb-1">
                    {currentHabit.topBarrier.label}
                  </p>
                  <p className="text-xs text-slate-600">
                    Memiliki skor terendah (<strong>{currentHabit.topBarrier.mean}</strong>) pada
                    dimensi ini dan menjadi titik utama penyumbang penurunan skor siswa.
                  </p>
                </div>

                <div className="bg-emerald-50 border border-emerald-200 p-5 rounded-2xl">
                  <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm mb-2">
                    <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                    <span>Kekuatan Tertinggi ({currentHabit.topStrength.qCode})</span>
                  </div>
                  <p className="text-sm font-semibold text-slate-900 mb-1">
                    {currentHabit.topStrength.label}
                  </p>
                  <p className="text-xs text-slate-600">
                    Konsisten diterapkan mayoritas siswa dengan skor rata-rata{' '}
                    <strong>{currentHabit.topStrength.mean}</strong>.
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: PATTERNS & CORRELATION */}
        {activeTab === 'patterns' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            {/* Stacked Distribution Chart */}
            <div className="lg:col-span-6 bg-white p-6 rounded-2xl shadow-sm border border-slate-200/80">
              <h2 className="text-base font-bold text-slate-900">
                Distribusi Frekuensi Jawaban per Kebiasaan (%)
              </h2>
              <p className="text-xs text-slate-500 mb-4">
                Proporsi jawaban dari Skala 1 (Sangat Jarang) hingga Skala 5 (Sangat Sering)
              </p>
              <div className="h-80">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={analytics.habits.map((h) => ({
                      name: h.shortName,
                      'Sangat Jarang (1)': h.distPct[0],
                      'Jarang (2)': h.distPct[1],
                      'Kadang-kadang (3)': h.distPct[2],
                      'Sering (4)': h.distPct[3],
                      'Sangat Sering (5)': h.distPct[4],
                    }))}
                    layout="vertical"
                    margin={{ top: 5, right: 20, left: 25, bottom: 5 }}
  
