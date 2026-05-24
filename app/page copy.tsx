"use client"

import { useEffect, useMemo, useState } from "react"
import { supabase } from "@/lib/supabase"
import jsPDF from "jspdf"
import autoTable from "jspdf-autotable"
import * as XLSX from "xlsx"
import { openJobPrint } from "./pdfTemplates"

type Role = "admin" | "staff"
type Tab = "panel" | "newJob" | "customers" | "allJobs" | "invoice" | "reports" | "archive" | "stocks" | "deliveries" | "costs"
type Status = "printing" | "cover" | "delivery" | "finished"
type InvoiceStatus = "waiting" | "invoiced" | "paid"
type Priority = "normal" | "urgent"

type Profile = { id: string; email: string; full_name: string; role: Role }
type Customer = { id: number; company: string; person: string; phone: string; email: string; address: string; note: string; created_at?: string }
type Log = { id?: number; job_id: number; text: string; created_at: string }
type Delivery = { id: number; job_id: number; amount: number; note: string; created_at: string }
type Stock = { id: number; name: string; type: string; quantity: number; unit: string; min_quantity: number; created_at?: string }

type Job = {
  id: number
  customer_id: number
  customer_name: string
  job_name: string
  quantity: number
  delivered: number
  price: number
  paper_cost: number
  print_cost: number
  binding_cost: number
  lamination_cost: number
  labor_cost: number
  total_cost: number
  profit: number
  status: Status
  invoice_status: InvoiceStatus
  priority: Priority
  archived: boolean
  deadline: string
  note: string
  shrink_amount: string
  size: string
  page_count: string
  color: string
  print_type: string
  binding: string
  lamination: string
  cellophane: string
  inner_paper_gram: string
  inner_paper_type: string
  inner_paper_size: string
  inner_paper_amount: string
  cover_gram: string
  cover_type: string
  cover_paper_type: string
  cover_size: string
  cover_amount: string
  cover_inside_print: string
  created_at: string
  logs?: Log[]
  deliveries?: Delivery[]
}

const statusTitle: Record<Status, string> = { printing: "Baskıda", cover: "Kapak Takma", delivery: "Teslimat", finished: "Biten İşler" }
const invoiceTitle: Record<InvoiceStatus, string> = { waiting: "Fatura Bekliyor", invoiced: "Fatura Kesildi", paid: "Ödendi" }

const emptyCustomer = { company: "", person: "", phone: "", email: "", address: "", note: "" }
const emptyStock = { name: "", type: "Kağıt", quantity: "", unit: "Tabaka", min_quantity: "" }
const emptyJob = {
  customer_id: 0,
  job_name: "",
  quantity: "",
  price: "",
  paper_cost: "",
  print_cost: "",
  binding_cost: "",
  lamination_cost: "",
  labor_cost: "",
  deadline: "",
  priority: "normal" as Priority,
  shrink_amount: "",
  size: "",
  page_count: "",
  color: "4+4",
  print_type: "",
  binding: "",
  lamination: "Yok",
  cellophane: "Yok",
  inner_paper_gram: "",
  inner_paper_type: "",
  inner_paper_size: "",
  inner_paper_amount: "",
  cover_gram: "",
  cover_type: "",
  cover_paper_type: "",
  cover_size: "",
  cover_amount: "",
  cover_inside_print: "Yok",
  note: "",
}

export default function Home() {
  const [user, setUser] = useState<any>(null)
  const [profile, setProfile] = useState<Profile | null>(null)
  const [authMode, setAuthMode] = useState<"login" | "register">("login")
  const [auth, setAuth] = useState({ email: "", password: "", fullName: "" })
  const [loading, setLoading] = useState(true)
  const [tab, setTab] = useState<Tab>("panel")
  const [search, setSearch] = useState("")
  const [monthFilter, setMonthFilter] = useState(new Date().toISOString().slice(0, 7))
  const [customers, setCustomers] = useState<Customer[]>([])
  const [jobs, setJobs] = useState<Job[]>([])
  const [stocks, setStocks] = useState<Stock[]>([])
  const [customerForm, setCustomerForm] = useState(emptyCustomer)
  const [stockForm, setStockForm] = useState(emptyStock)
  const [jobForm, setJobForm] = useState(emptyJob)
  const [editJob, setEditJob] = useState<Job | null>(null)

  const isAdmin = profile?.role === "admin"
  const activeJobs = jobs.filter(j => !j.archived)
  const archivedJobs = jobs.filter(j => j.archived)

  const stats = useMemo(() => ({
    total: activeJobs.length,
    urgent: activeJobs.filter(j => j.priority === "urgent").length,
    late: activeJobs.filter(j => isLate(j)).length,
    printing: activeJobs.filter(j => j.status === "printing").length,
    cover: activeJobs.filter(j => j.status === "cover").length,
    delivery: activeJobs.filter(j => j.status === "delivery").length,
    finished: activeJobs.filter(j => j.status === "finished").length,
    revenue: activeJobs.reduce((s, j) => s + Number(j.price || 0), 0),
    profit: activeJobs.reduce((s, j) => s + Number(j.profit || 0), 0),
    stockWarning: stocks.filter(s => Number(s.quantity) <= Number(s.min_quantity)).length,
  }), [jobs, stocks])
  const today = new Date().toISOString().slice(0, 10)

const alerts = activeJobs
  .filter((j) => {
    const isUrgent = j.priority === "urgent"
    const isToday = j.deadline === today
    const isLate = j.deadline && j.deadline < today && j.status !== "finished"

    return isUrgent || isToday || isLate
  })
  .map((j) => {
    const isLate = j.deadline && j.deadline < today && j.status !== "finished"
    const isToday = j.deadline === today

    return {
      id: j.id,
      title: `${jobNo(j)} - ${j.customer_name}`,
      text: isLate
        ? "Teslim tarihi geçmiş"
        : isToday
        ? "Bugün teslim edilecek"
        : "Acil iş",
      color: isLate
        ? "bg-red-100 text-red-700"
        : isToday
        ? "bg-yellow-100 text-yellow-700"
        : "bg-orange-100 text-orange-700",
    }
  })

  const visibleTabs: { key: Tab; text: string }[] = [
    { key: "panel", text: "İş Takip Paneli" },
    ...(isAdmin ? [{ key: "newJob" as Tab, text: "Yeni İş Girişi" }] : []),
    ...(isAdmin ? [{ key: "customers" as Tab, text: "Müşteriler" }] : []),
    { key: "allJobs", text: "Tüm İşler" },
    ...(isAdmin ? [
      { key: "invoice" as Tab, text: "Faturalar" },
      { key: "costs" as Tab, text: "Maliyet / Kâr" },
      { key: "stocks" as Tab, text: "Stok" },
      { key: "deliveries" as Tab, text: "Teslimat Geçmişi" },
      { key: "reports" as Tab, text: "Raporlar" },
      { key: "archive" as Tab, text: "Arşiv" },
    ] : []),
  ]

  useEffect(() => {
    supabase.auth.getUser().then(({ data }) => {
      setUser(data.user)
      if (data.user) init(data.user)
      else setLoading(false)
    })
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user || null)
      if (session?.user) init(session.user)
    })
    return () => sub.subscription.unsubscribe()
  }, [])

  async function init(currentUser: any) {
    setLoading(true)
    await ensureProfile(currentUser)
    await loadAll()
    setLoading(false)
  }

  async function ensureProfile(currentUser: any) {
    const { data: existing } = await supabase.from("profiles").select("*").eq("id", currentUser.id).maybeSingle()
    if (existing) { setProfile(existing as Profile); return }
    const { count } = await supabase.from("profiles").select("*", { count: "exact", head: true })
    const newProfile = {
      id: currentUser.id,
      email: currentUser.email,
      full_name: currentUser.user_metadata?.full_name || currentUser.email,
      role: count === 0 ? "admin" : "staff",
    }
    await supabase.from("profiles").insert(newProfile)
    setProfile(newProfile as Profile)
  }

  async function loadAll() {
    const [{ data: cs }, { data: js }, { data: ls }, { data: ds }, { data: st }] = await Promise.all([
      supabase.from("customers").select("*").order("id", { ascending: false }),
      supabase.from("jobs").select("*").order("id", { ascending: false }),
      supabase.from("logs").select("*").order("id", { ascending: false }),
      supabase.from("deliveries").select("*").order("id", { ascending: false }),
      supabase.from("stocks").select("*").order("id", { ascending: false }),
    ])
    const jobsWith = ((js || []) as Job[]).map(j => ({ ...j, logs: (ls || []).filter((l: any) => l.job_id === j.id), deliveries: (ds || []).filter((d: any) => d.job_id === j.id) }))
    setCustomers((cs || []) as Customer[])
    setJobs(jobsWith)
    setStocks((st || []) as Stock[])
    if ((cs || [])[0]) setJobForm(p => ({ ...p, customer_id: (cs || [])[0].id }))
  }

  async function login() {
    const { error } = await supabase.auth.signInWithPassword({ email: auth.email, password: auth.password })
    if (error) alert(error.message)
  }

  async function register() {
    const { error } = await supabase.auth.signUp({ email: auth.email, password: auth.password, options: { data: { full_name: auth.fullName } } })
    if (error) alert(error.message)
    else alert("Kayıt oluşturuldu. Giriş yapabilirsin. E-posta doğrulama açıksa mailini kontrol et.")
  }

  async function logout() { await supabase.auth.signOut(); setProfile(null); setUser(null) }
  async function addLog(jobId: number, text: string) { await supabase.from("logs").insert({ job_id: jobId, text }) }

  async function addCustomer() {
    if (!customerForm.company) return alert("Firma adı zorunlu.")
    const { error } = await supabase.from("customers").insert(customerForm)
    if (error) return alert(error.message)
    setCustomerForm(emptyCustomer); await loadAll()
  }

  function calcTotals(data: typeof emptyJob) {
    const total_cost = [data.paper_cost, data.print_cost, data.binding_cost, data.lamination_cost, data.labor_cost].reduce((s, v) => s + Number(v || 0), 0)
    const profit = Number(data.price || 0) - total_cost
    return { total_cost, profit }
  }

  async function addJob() {
    const customer = customers.find(c => c.id === Number(jobForm.customer_id))
    if (!customer || !jobForm.job_name) return alert("Müşteri ve iş adı zorunlu.")
    const totals = calcTotals(jobForm)
    const payload = {
      customer_id: customer.id,
      customer_name: customer.company,
      job_name: jobForm.job_name,
      quantity: Number(jobForm.quantity || 0), delivered: 0,
      price: Number(jobForm.price || 0),
      paper_cost: Number(jobForm.paper_cost || 0), print_cost: Number(jobForm.print_cost || 0), binding_cost: Number(jobForm.binding_cost || 0), lamination_cost: Number(jobForm.lamination_cost || 0), labor_cost: Number(jobForm.labor_cost || 0),
      total_cost: totals.total_cost, profit: totals.profit,
      status: "printing", invoice_status: "waiting", priority: jobForm.priority, archived: false,
      deadline: jobForm.deadline, note: jobForm.note,
      shrink_amount: jobForm.shrink_amount, size: jobForm.size, page_count: jobForm.page_count, color: jobForm.color, print_type: jobForm.print_type, binding: jobForm.binding, lamination: jobForm.lamination, cellophane: jobForm.cellophane,
      inner_paper_gram: jobForm.inner_paper_gram, inner_paper_type: jobForm.inner_paper_type, inner_paper_size: jobForm.inner_paper_size, inner_paper_amount: jobForm.inner_paper_amount,
      cover_gram: jobForm.cover_gram, cover_type: jobForm.cover_type, cover_paper_type: jobForm.cover_paper_type, cover_size: jobForm.cover_size, cover_amount: jobForm.cover_amount, cover_inside_print: jobForm.cover_inside_print,
    }
    const { data, error } = await supabase.from("jobs").insert(payload).select().single()
    if (error) return alert(error.message)
    await addLog(data.id, "İş oluşturuldu.")
    setJobForm({ ...emptyJob, customer_id: customers[0]?.id || 0 })
    setTab("panel"); await loadAll()
  }

  async function saveEditJob() {
    if (!editJob) return
    const total_cost = Number(editJob.paper_cost || 0) + Number(editJob.print_cost || 0) + Number(editJob.binding_cost || 0) + Number(editJob.lamination_cost || 0) + Number(editJob.labor_cost || 0)
    const profit = Number(editJob.price || 0) - total_cost
    const { logs, deliveries, ...clean } = editJob
    await supabase.from("jobs").update({ ...clean, total_cost, profit }).eq("id", editJob.id)
    await addLog(editJob.id, "İş bilgileri düzenlendi.")
    setEditJob(null); await loadAll()
  }

  async function updateJob(job: Job, updates: Partial<Job>, logText: string) {
    await supabase.from("jobs").update(updates).eq("id", job.id)
    await addLog(job.id, logText); await loadAll()
  }

  async function nextJob(job: Job) {
    if (job.status === "printing") return updateJob(job, { status: "cover" }, "Baskı tamamlandı, kapak takmaya geçti.")
    if (job.status === "cover") return updateJob(job, { status: "delivery" }, "Kapak takıldı, teslimata geçti.")
    if (job.status === "delivery") {
      const amount = Number(prompt("Kaç adet teslim edildi?") || 0)
      if (amount <= 0) return
      const note = prompt("Teslimat notu") || ""
      const delivered = Number(job.delivered || 0) + amount
      await supabase.from("deliveries").insert({ job_id: job.id, amount, note })
      await updateJob(job, { delivered, status: delivered >= job.quantity ? "finished" : "delivery" }, `${amount} adet teslim edildi.`)
    }
  }

  async function prevJob(job: Job) {
    if (job.status === "cover") return updateJob(job, { status: "printing" }, "Geri alındı: Baskıya döndü.")
    if (job.status === "delivery") return updateJob(job, { status: "cover" }, "Geri alındı: Kapak takmaya döndü.")
    if (job.status === "finished") return updateJob(job, { status: "delivery" }, "Geri alındı: Teslimata döndü.")
  }

  async function archiveJob(job: Job) { if (confirm("Arşive taşınsın mı?")) await updateJob(job, { archived: true }, "İş arşive taşındı.") }
  async function restoreJob(job: Job) { await updateJob(job, { archived: false }, "İş arşivden geri alındı.") }
  async function updateInvoice(job: Job, invoice_status: InvoiceStatus) { await updateJob(job, { invoice_status }, `Fatura durumu: ${invoiceTitle[invoice_status]}`) }

  async function copyJob(job: Job) {
    const { id, logs, deliveries, created_at, ...copy } = job
    const { data } = await supabase.from("jobs").insert({ ...copy, job_name: `${job.job_name} - Kopya`, delivered: 0, status: "printing", invoice_status: "waiting", archived: false }).select().single()
    if (data) await addLog(data.id, `{jobNo(job)} üzerinden kopyalandı.`)
    await loadAll()
  }

  async function addStock() {
    if (!stockForm.name) return alert("Stok adı zorunlu.")
    await supabase.from("stocks").insert({ name: stockForm.name, type: stockForm.type, quantity: Number(stockForm.quantity || 0), unit: stockForm.unit, min_quantity: Number(stockForm.min_quantity || 0) })
    setStockForm(emptyStock); await loadAll()
  }

  async function moveStock(stock: Stock, movement_type: "Giriş" | "Çıkış") {
    const amount = Number(prompt(`${movement_type} miktarı`) || 0)
    if (amount <= 0) return
    const newQty = movement_type === "Giriş" ? Number(stock.quantity) + amount : Number(stock.quantity) - amount
    await supabase.from("stocks").update({ quantity: newQty }).eq("id", stock.id)
    await supabase.from("stock_movements").insert({ stock_id: stock.id, movement_type, quantity: amount, note: `${movement_type} yapıldı` })
    await loadAll()
  }

  function openPdf(pdf: jsPDF) { const url = URL.createObjectURL(pdf.output("blob")); window.open(url, "_blank") }
  function pdfHeader(pdf: jsPDF, title: string, sub: string) {
    pdf.setFillColor(7, 29, 53); pdf.rect(0, 0, 210, 30, "F")
    pdf.setTextColor(255,255,255); pdf.setFontSize(20); pdf.text("BAL BASIM YAYIN", 14, 13); pdf.setFontSize(12); pdf.text(title, 14, 23); pdf.text(sub, 160, 20); pdf.setTextColor(0,0,0)
  }
  function table(pdf: jsPDF, startY: number, head: string[][], body: any[][]) { autoTable(pdf, { startY, head, body, theme: "grid", headStyles: { fillColor: [7,29,53], textColor: 255, fontStyle: "bold" }, alternateRowStyles: { fillColor: [245,247,251] }, styles: { fontSize: 9, cellPadding: 3, lineColor: [210,210,210] } }) }
function makeJobPdf(job: Job) {
  openJobPrint(job)
}
  function makeDeliveryPdf(job: Job) {
    const pdf = new jsPDF(); pdfHeader(pdf, "Teslim Fişi", `{jobNo(job)}`)
    const last = job.deliveries?.[0]?.amount || "-"
    table(pdf, 45, [["Alan", "Bilgi"]], [["Firma", job.customer_name], ["İş", job.job_name], ["Toplam Adet", job.quantity], ["Son Teslim", last], ["Toplam Teslim", job.delivered], ["Kalan", Math.max(job.quantity - job.delivered, 0)], ["Tarih", new Date().toLocaleDateString("tr-TR")]])
    pdf.text("Teslim Eden", 25, 230); pdf.text("Teslim Alan", 135, 230); pdf.line(25,240,75,240); pdf.line(135,240,185,240); openPdf(pdf)
  }
  function makeMonthlyReportPdf(monthJobs: Job[]) {
    const pdf = new jsPDF(); const total = monthJobs.reduce((s,j)=>s+Number(j.price||0),0); const cost = monthJobs.reduce((s,j)=>s+Number(j.total_cost||0),0); const profit = total-cost
    pdfHeader(pdf, "Aylık ERP Raporu", monthFilter)
    table(pdf, 40, [["Özet", "Bilgi"]], [["Toplam İş", monthJobs.length], ["Toplam Ciro", `${total.toLocaleString("tr-TR")} TL`], ["Toplam Maliyet", `${cost.toLocaleString("tr-TR")} TL`], ["Toplam Kâr", `${profit.toLocaleString("tr-TR")} TL`]])
    table(pdf, 95, [["İş No", "Müşteri", "İş", "Adet", "Ciro", "Maliyet", "Kâr"]], monthJobs.map(j=>[`İŞ-${j.id}`, j.customer_name, j.job_name, j.quantity, `${j.price} TL`, `${j.total_cost} TL`, `${j.profit} TL`]))
    openPdf(pdf)
  }

  function exportJobsExcel() {
    const rows = activeJobs.map(j => ({ İşNo: `İŞ-${j.id}`, Müşteri: j.customer_name, İş: j.job_name, Adet: j.quantity, Teslim: j.delivered, Durum: statusTitle[j.status], Fatura: invoiceTitle[j.invoice_status], Fiyat: j.price, Maliyet: j.total_cost, Kar: j.profit }))
    const wb = XLSX.utils.book_new(); const ws = XLSX.utils.json_to_sheet(rows); XLSX.utils.book_append_sheet(wb, ws, "İşler"); XLSX.writeFile(wb, "bal-basim-isler.xlsx")
  }

  function isLate(job: Job) { if (!job.deadline || job.status === "finished") return false; return new Date(job.deadline) < new Date(new Date().toDateString()) }
  const searchedJobs = activeJobs.filter(j => `${j.id} ${j.customer_name} ${j.job_name} ${statusTitle[j.status]}`.toLowerCase().includes(search.toLowerCase()))
  const monthlyJobs = activeJobs.filter(j => j.created_at?.slice(0,7) === monthFilter)

  if (loading) return <div className="p-10 text-2xl font-black">Yükleniyor...</div>
  if (!user) return <Login auth={auth} setAuth={setAuth} authMode={authMode} setAuthMode={setAuthMode} login={login} register={register} />

  return <main className="min-h-screen bg-[#f4f7fb] text-slate-900">
    <aside className="fixed left-0 top-0 h-screen overflow-y-auto w-[270px] bg-[#071d35] text-white p-5">
      <div className="mb-8"><img src="/logo.png" className="w-40 mb-3" /><div className="text-sm text-slate-300">Premium ERP</div></div>
      <div className="mb-5 text-sm"><b>{profile?.full_name}</b><br/><span className="text-slate-300">{profile?.role === "admin" ? "Yönetici" : "Personel"}</span></div>
      {visibleTabs.map(i => <button key={i.key} onClick={()=>setTab(i.key)} className={`w-full text-left px-4 py-3 rounded-lg mb-2 text-sm ${tab===i.key ? "bg-blue-600" : "text-slate-200 hover:bg-slate-800"}`}>{i.text}</button>)}
      <button
  onClick={logout}
  className="w-full mt-6 bg-red-600 py-3 rounded-lg font-bold"
>Çıkış</button>
    </aside>
    <section className="ml-[270px] p-7"><h1 className="text-2xl font-black mb-6">{tabTitle(tab)}</h1>
      {tab==="panel" && <><div className="grid grid-cols-8 gap-4 mb-7"><Stat title="Toplam" value={stats.total}/><Stat title="Acil" value={stats.urgent}/><Stat title="Geciken" value={stats.late}/><Stat title="Baskı" value={stats.printing}/><Stat title="Kapak" value={stats.cover}/><Stat title="Teslim" value={stats.delivery}/><Stat title="Biten" value={stats.finished}/>{isAdmin && <Stat title="Kâr" valueText={`${stats.profit.toLocaleString("tr-TR")} ₺`}/>}</div><div className="grid grid-cols-4 gap-5">{(["printing","cover","delivery","finished"] as Status[]).map(s => <Column key={s} status={s} jobs={activeJobs.filter(j=>j.status===s)} isAdmin={isAdmin} nextJob={nextJob} prevJob={prevJob} archiveJob={archiveJob} copyJob={copyJob} setEditJob={setEditJob} makeJobPdf={makeJobPdf} makeDeliveryPdf={makeDeliveryPdf}/>)}</div></>}
      {tab==="newJob" && isAdmin && <JobForm customers={customers} jobForm={jobForm} setJobForm={setJobForm} addJob={addJob} />}
      {tab==="customers" && isAdmin && <Customers customers={customers} customerForm={customerForm} setCustomerForm={setCustomerForm} addCustomer={addCustomer}/>} 
      {tab==="allJobs" && <Panel><div className="flex justify-between mb-4"><input className="border rounded-lg p-2 w-[420px]" placeholder="Ara..." value={search} onChange={e=>setSearch(e.target.value)}/>{isAdmin && <button onClick={exportJobsExcel} className="bg-green-600 text-white px-4 py-2 rounded-lg font-bold">Excel Aktar</button>}</div><JobsTable jobs={searchedJobs} isAdmin={isAdmin} makeJobPdf={makeJobPdf}/></Panel>}
      {tab==="invoice" && isAdmin && <Panel><InvoiceTable jobs={activeJobs.filter(j=>j.status==="finished")} updateInvoice={updateInvoice} makeJobPdf={makeJobPdf}/></Panel>}
      {tab==="costs" && isAdmin && <Panel><JobsTable jobs={activeJobs} isAdmin={true} makeJobPdf={makeJobPdf} showCosts /></Panel>}
      {tab==="stocks" && isAdmin && <Stocks stocks={stocks} stockForm={stockForm} setStockForm={setStockForm} addStock={addStock} moveStock={moveStock}/>} 
      {tab==="deliveries" && isAdmin && <Panel><Deliveries jobs={activeJobs}/></Panel>}
      {tab==="reports" && isAdmin && <Panel><div className="flex justify-between mb-5"><input type="month" className="border rounded-lg p-2" value={monthFilter} onChange={e=>setMonthFilter(e.target.value)}/><button onClick={()=>makeMonthlyReportPdf(monthlyJobs)} className="bg-slate-900 text-white px-4 py-2 rounded-lg font-bold">Aylık PDF</button></div><JobsTable jobs={monthlyJobs} isAdmin={true} makeJobPdf={makeJobPdf} showCosts /></Panel>}
      {tab==="archive" && isAdmin && <Panel>{archivedJobs.map(j=><div key={j.id} className="border rounded-lg p-3 mb-2 flex justify-between"><b>{jobNo(j)} {j.customer_name} - {j.job_name}</b><button onClick={()=>restoreJob(j)} className="bg-blue-600 text-white px-3 py-2 rounded">Geri Al</button></div>)}</Panel>}
    {alerts.length > 0 && (
  <div className="bg-white border rounded-xl p-5 mb-6">
    <h2 className="font-black text-lg mb-3">🔔 Canlı Bildirimler</h2>

    <div className="grid grid-cols-3 gap-3">
      {alerts.map((a) => (
        <div key={a.id} className={`rounded-lg p-3 font-bold ${a.color}`}>
          <div>{a.title}</div>
          <div className="text-sm opacity-80">{a.text}</div>
        </div>
      ))}
    </div>
  </div>
)}
    </section>
    {editJob && <EditModal job={editJob} setJob={setEditJob} save={saveEditJob} />}
  </main>
}

function Login({ auth, setAuth, authMode, setAuthMode, login, register }: any) { return <main className="min-h-screen bg-[#071d35] flex items-center justify-center"><div className="bg-white rounded-2xl p-8 w-[420px]"><img src="/logo.png" className="w-52 mx-auto mb-4" /><p className="text-slate-500 mb-6">Premium ERP Giriş</p>{authMode==="register" && <Input p="Ad Soyad" v={auth.fullName} c={(v:string)=>setAuth({...auth, fullName:v})}/>}<div className="mt-3"><Input p="E-posta" v={auth.email} c={(v:string)=>setAuth({...auth,email:v})}/></div><div className="mt-3"><Input p="Şifre" type="password" v={auth.password} c={(v:string)=>setAuth({...auth,password:v})}/></div><button onClick={authMode==="login"?login:register} className="w-full mt-5 bg-blue-600 text-white py-3 rounded-lg font-bold">{authMode==="login"?"Giriş Yap":"Kayıt Ol"}</button><button onClick={()=>setAuthMode(authMode==="login"?"register":"login")} className="w-full mt-3 text-blue-600">{authMode==="login"?"Hesap oluştur":"Girişe dön"}</button></div></main> }
function jobNo(job: Job) {
  const year = String(new Date(job.created_at).getFullYear()).slice(2)

  return `${String(job.id).padStart(4, "0")}-${year}`
}
function tabTitle(tab: Tab) { return { panel:"📋 İş Takip Paneli", newJob:"➕ Yeni İş Girişi", customers:"👥 Müşteriler", allJobs:"📚 Tüm İşler", invoice:"💰 Faturalar", reports:"📊 Raporlar", archive:"🗄 Arşiv", stocks:"📦 Stok", deliveries:"🚚 Teslimat Geçmişi", costs:"📈 Maliyet / Kâr" }[tab] }
function Panel({children}: any){ return <div className="bg-white border rounded-xl p-5">{children}</div> }
function Section({ title, children }: any){ return <Panel><h2 className="font-black text-xl mb-4">{title}</h2><div className="grid grid-cols-4 gap-3">{children}</div></Panel> }
function Input({ p, v, c, type="text" }: any){ return <input className="border rounded-lg p-2 text-sm w-full" placeholder={p} value={v||""} type={type} onChange={e=>c(e.target.value)}/> }
function RadioGroup({label,value,options,onChange}: any){ return <div><label className="text-xs font-bold text-slate-600">{label}</label><div className="flex gap-2 mt-1">{options.map((o:any)=><button key={o.value} type="button" onClick={()=>onChange(o.value)} className={`px-4 py-2 rounded-lg border text-sm font-bold ${value===o.value?"bg-blue-600 text-white":"bg-white"}`}>{o.label}</button>)}</div></div> }
function Stat({title,value,valueText}: any){ return <div className="bg-white border rounded-xl p-5"><div className="text-sm text-slate-500 font-bold">{title}</div><div className="text-2xl font-black mt-2">{valueText||value}</div></div> }
function JobForm({customers,jobForm,setJobForm,addJob}: any){ return <div className="space-y-5"><Section title="Genel İş Bilgileri"><select className="border rounded-lg p-2 text-sm" value={jobForm.customer_id} onChange={e=>setJobForm({...jobForm, customer_id:Number(e.target.value)})}>{customers.map((c:Customer)=><option key={c.id} value={c.id}>{c.company}</option>)}</select><Input p="İş Adı" v={jobForm.job_name} c={(v:string)=>setJobForm({...jobForm,job_name:v})}/><Input p="Adet" type="number" v={jobForm.quantity} c={(v:string)=>setJobForm({...jobForm,quantity:v})}/><Input p="Fiyat" type="number" v={jobForm.price} c={(v:string)=>setJobForm({...jobForm,price:v})}/><Input p="Teslim Tarihi" type="date" v={jobForm.deadline} c={(v:string)=>setJobForm({...jobForm,deadline:v})}/><RadioGroup label="Öncelik" value={jobForm.priority} options={[{label:"Normal",value:"normal"},{label:"Acil",value:"urgent"}]} onChange={(v:string)=>setJobForm({...jobForm,priority:v})}/></Section><Section title="Maliyet Bilgileri"><Input p="Kağıt Maliyeti" type="number" v={jobForm.paper_cost} c={(v:string)=>setJobForm({...jobForm,paper_cost:v})}/><Input p="Baskı Maliyeti" type="number" v={jobForm.print_cost} c={(v:string)=>setJobForm({...jobForm,print_cost:v})}/><Input p="Cilt Maliyeti" type="number" v={jobForm.binding_cost} c={(v:string)=>setJobForm({...jobForm,binding_cost:v})}/><Input p="Laminasyon Maliyeti" type="number" v={jobForm.lamination_cost} c={(v:string)=>setJobForm({...jobForm,lamination_cost:v})}/><Input p="İşçilik" type="number" v={jobForm.labor_cost} c={(v:string)=>setJobForm({...jobForm,labor_cost:v})}/></Section><Section title="Baskı Bilgileri"><Input p="Shrink Adedi" v={jobForm.shrink_amount} c={(v:string)=>setJobForm({...jobForm,shrink_amount:v})}/><Input p="Ebat" v={jobForm.size} c={(v:string)=>setJobForm({...jobForm,size:v})}/><Input p="Sayfa Sayısı" v={jobForm.page_count} c={(v:string)=>setJobForm({...jobForm,page_count:v})}/><Input p="Renk" v={jobForm.color} c={(v:string)=>setJobForm({...jobForm,color:v})}/><Input p="Baskı Tipi" v={jobForm.print_type} c={(v:string)=>setJobForm({...jobForm,print_type:v})}/><Input p="Cilt Şekli" v={jobForm.binding} c={(v:string)=>setJobForm({...jobForm,binding:v})}/><RadioGroup label="Laminasyon" value={jobForm.lamination} options={[{label:"Yok",value:"Yok"},{label:"Var",value:"Var"}]} onChange={(v:string)=>setJobForm({...jobForm,lamination:v})}/><RadioGroup label="Selefon" value={jobForm.cellophane} options={[{label:"Yok",value:"Yok"},{label:"Mat",value:"Mat"},{label:"Parlak",value:"Parlak"}]} onChange={(v:string)=>setJobForm({...jobForm,cellophane:v})}/></Section><Section title="İç Kağıt Bilgileri"><Input p="Gramaj" v={jobForm.inner_paper_gram} c={(v:string)=>setJobForm({...jobForm,inner_paper_gram:v})}/><Input p="Türü" v={jobForm.inner_paper_type} c={(v:string)=>setJobForm({...jobForm,inner_paper_type:v})}/><Input p="Ebatı" v={jobForm.inner_paper_size} c={(v:string)=>setJobForm({...jobForm,inner_paper_size:v})}/><Input p="Miktarı" v={jobForm.inner_paper_amount} c={(v:string)=>setJobForm({...jobForm,inner_paper_amount:v})}/></Section><Section title="Kapak Kağıdı Bilgileri"><Input p="Gramaj" v={jobForm.cover_gram} c={(v:string)=>setJobForm({...jobForm,cover_gram:v})}/><Input p="Türü" v={jobForm.cover_type} c={(v:string)=>setJobForm({...jobForm,cover_type:v})}/><Input p="Kağıt Türü" v={jobForm.cover_paper_type} c={(v:string)=>setJobForm({...jobForm,cover_paper_type:v})}/><Input p="Ebat" v={jobForm.cover_size} c={(v:string)=>setJobForm({...jobForm,cover_size:v})}/><Input p="Miktar" v={jobForm.cover_amount} c={(v:string)=>setJobForm({...jobForm,cover_amount:v})}/><RadioGroup label="Kapak İçi Baskı" value={jobForm.cover_inside_print} options={[{label:"Yok",value:"Yok"},{label:"Var",value:"Var"}]} onChange={(v:string)=>setJobForm({...jobForm,cover_inside_print:v})}/></Section><textarea className="border rounded-lg p-3 w-full" placeholder="Not" value={jobForm.note} onChange={e=>setJobForm({...jobForm,note:e.target.value})}/><button onClick={addJob} className="bg-blue-600 text-white px-6 py-3 rounded-lg font-bold">İşi Kaydet</button></div> }
function Customers({customers,customerForm,setCustomerForm,addCustomer}: any){ return <div className="grid grid-cols-3 gap-5"><Panel><h2 className="font-black text-xl mb-4">Müşteri Kayıt</h2><div className="space-y-3"><Input p="Firma" v={customerForm.company} c={(v:string)=>setCustomerForm({...customerForm,company:v})}/><Input p="Yetkili" v={customerForm.person} c={(v:string)=>setCustomerForm({...customerForm,person:v})}/><Input p="Telefon" v={customerForm.phone} c={(v:string)=>setCustomerForm({...customerForm,phone:v})}/><Input p="E-posta" v={customerForm.email} c={(v:string)=>setCustomerForm({...customerForm,email:v})}/><Input p="Adres" v={customerForm.address} c={(v:string)=>setCustomerForm({...customerForm,address:v})}/><Input p="Not" v={customerForm.note} c={(v:string)=>setCustomerForm({...customerForm,note:v})}/><button onClick={addCustomer} className="bg-blue-600 text-white px-5 py-3 rounded-lg font-bold">Müşteri Ekle</button></div></Panel><div className="col-span-2"><Panel><SimpleTable head={["Firma","Yetkili","Telefon","Adres","Not"]} rows={customers.map((c:Customer)=>[c.company,c.person,c.phone,c.address,c.note])}/></Panel></div></div> }
function Column({status,jobs,isAdmin,nextJob,prevJob,archiveJob,copyJob,setEditJob,makeJobPdf,makeDeliveryPdf}: any){ return <div className="bg-white border border-t-4 border-t-blue-600 rounded-xl p-3 min-h-[560px]"><div className="flex justify-between mb-3"><h3 className="font-black text-sm">{statusTitle[status as Status]}</h3><span>{jobs.length}</span></div><div className="space-y-3">{jobs.map((job:Job)=><div key={job.id} className={`border rounded-lg p-3 shadow-sm ${job.priority==="urgent"?"bg-red-50 border-red-300":"bg-white"}`}><button onClick={()=>makeJobPdf(job)} className="text-xs text-blue-600 font-bold hover:underline">{jobNo(job)}</button><div className="font-black mt-2">{job.customer_name}</div><div className="text-sm">{job.job_name}</div><div className="text-sm mt-2">Adet: {job.quantity}</div><div className="text-sm">Teslim: {job.delivered}</div><div className="text-sm">Kalan: {Math.max(job.quantity-job.delivered,0)}</div>{job.priority==="urgent"&&<div className="text-red-600 text-xs font-black">ACİL</div>}{isAdmin&&<div className="text-sm font-bold">Kâr: {Number(job.profit||0).toLocaleString("tr-TR")} ₺</div>}<div className="grid grid-cols-2 gap-2 mt-3">{status!=="printing"&&<button onClick={()=>prevJob(job)} className="bg-slate-200 py-2 rounded text-xs font-bold">← Geri</button>}{status!=="finished"&&<button onClick={()=>nextJob(job)} className="bg-blue-600 text-white py-2 rounded text-xs font-bold">İleri →</button>}</div>{job.delivered>0&&<button onClick={()=>makeDeliveryPdf(job)} className="w-full mt-2 bg-green-600 text-white py-2 rounded text-xs font-bold">Teslim Fişi</button>}{isAdmin&&<><button onClick={()=>setEditJob(job)} className="w-full mt-2 bg-yellow-100 text-yellow-800 py-2 rounded text-xs font-bold">Düzenle</button><button onClick={()=>copyJob(job)} className="w-full mt-2 bg-slate-900 text-white py-2 rounded text-xs font-bold">Kopyala</button><button onClick={()=>archiveJob(job)} className="w-full mt-2 bg-red-100 text-red-700 py-2 rounded text-xs font-bold">Arşive Taşı</button></>}<details className="mt-2 text-xs"><summary className="cursor-pointer font-bold">Geçmiş</summary>{(job.logs||[]).map((l,i)=><div key={i} className="border-t py-1">{new Date(l.created_at).toLocaleString("tr-TR")} - {l.text}</div>)}</details></div>)}</div></div> }
function JobsTable({jobs,isAdmin,makeJobPdf,showCosts}: any){ return <table className="w-full text-sm"><thead><tr className="border-b text-left text-slate-500"><th className="py-2">İş No</th><th>Müşteri</th><th>İş</th><th>Adet</th><th>Teslim</th>{isAdmin&&<th>Fiyat</th>}{showCosts&&<><th>Maliyet</th><th>Kâr</th></>}<th>Durum</th>{isAdmin&&<th>Fatura</th>}</tr></thead><tbody>{jobs.map((j:Job)=><tr key={j.id} className="border-b"><td className="py-3 font-bold"><button onClick={()=>makeJobPdf(j)} className="text-blue-600 hover:underline">{jobNo(j)}</button></td><td>{j.customer_name}</td><td>{j.job_name}</td><td>{j.quantity}</td><td>{j.delivered}</td>{isAdmin&&<td>{j.price.toLocaleString("tr-TR")} ₺</td>}{showCosts&&<><td>{Number(j.total_cost||0).toLocaleString("tr-TR")} ₺</td><td>{Number(j.profit||0).toLocaleString("tr-TR")} ₺</td></>}<td>{statusTitle[j.status]}</td>{isAdmin&&<td>{invoiceTitle[j.invoice_status]}</td>}</tr>)}</tbody></table> }
function InvoiceTable({jobs,updateInvoice,makeJobPdf}: any){ return <table className="w-full text-sm"><thead><tr className="border-b text-left text-slate-500"><th className="py-2">İş No</th><th>Müşteri</th><th>İş</th><th>Tutar</th><th>Durum</th><th>İşlem</th></tr></thead><tbody>{jobs.map((j:Job)=><tr key={j.id} className="border-b"><td className="py-3 font-bold"><button onClick={()=>makeJobPdf(j)} className="text-blue-600 hover:underline">{jobNo(j)}</button></td><td>{j.customer_name}</td><td>{j.job_name}</td><td>{j.price.toLocaleString("tr-TR")} ₺</td><td>{invoiceTitle[j.invoice_status]}</td><td className="space-x-2"><button onClick={()=>updateInvoice(j,"waiting")} className="bg-slate-200 px-3 py-2 rounded">Bekliyor</button><button onClick={()=>updateInvoice(j,"invoiced")} className="bg-green-600 text-white px-3 py-2 rounded">Kesildi</button><button onClick={()=>updateInvoice(j,"paid")} className="bg-blue-600 text-white px-3 py-2 rounded">Ödendi</button></td></tr>)}</tbody></table> }
function Stocks({stocks,stockForm,setStockForm,addStock,moveStock}: any){ return <div className="grid grid-cols-3 gap-5"><Panel><h2 className="font-black text-xl mb-4">Stok Ekle</h2><div className="space-y-3"><Input p="Stok Adı" v={stockForm.name} c={(v:string)=>setStockForm({...stockForm,name:v})}/><Input p="Tür" v={stockForm.type} c={(v:string)=>setStockForm({...stockForm,type:v})}/><Input p="Miktar" type="number" v={stockForm.quantity} c={(v:string)=>setStockForm({...stockForm,quantity:v})}/><Input p="Birim" v={stockForm.unit} c={(v:string)=>setStockForm({...stockForm,unit:v})}/><Input p="Minimum" type="number" v={stockForm.min_quantity} c={(v:string)=>setStockForm({...stockForm,min_quantity:v})}/><button onClick={addStock} className="bg-blue-600 text-white px-5 py-3 rounded-lg font-bold">Stok Ekle</button></div></Panel><div className="col-span-2"><Panel><table className="w-full text-sm"><thead><tr className="border-b text-left text-slate-500"><th>Ad</th><th>Tür</th><th>Miktar</th><th>Min</th><th>İşlem</th></tr></thead><tbody>{stocks.map((s:Stock)=><tr key={s.id} className="border-b"><td className="py-3 font-bold">{s.name}</td><td>{s.type}</td><td className={Number(s.quantity)<=Number(s.min_quantity)?"text-red-600 font-bold":""}>{s.quantity} {s.unit}</td><td>{s.min_quantity}</td><td className="space-x-2"><button onClick={()=>moveStock(s,"Giriş")} className="bg-green-600 text-white px-3 py-2 rounded">Giriş</button><button onClick={()=>moveStock(s,"Çıkış")} className="bg-red-600 text-white px-3 py-2 rounded">Çıkış</button></td></tr>)}</tbody></table></Panel></div></div> }
function Deliveries({jobs}: any){ const rows = jobs.flatMap((j:Job)=>(j.deliveries||[]).map(d=>[`İŞ-${j.id}`, j.customer_name, j.job_name, String(d.amount), d.note, new Date(d.created_at).toLocaleString("tr-TR")])) ; return <SimpleTable head={["İş No","Müşteri","İş","Adet","Not","Tarih"]} rows={rows}/> }
function EditModal({job,setJob,save}: any){ return <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50"><div className="bg-white rounded-xl p-6 w-[700px] max-h-[90vh] overflow-auto"><h2 className="text-xl font-black mb-4">İş Düzenle</h2><div className="grid grid-cols-2 gap-3"><Input p="İş Adı" v={job.job_name} c={(v:string)=>setJob({...job,job_name:v})}/><Input p="Adet" type="number" v={job.quantity} c={(v:string)=>setJob({...job,quantity:Number(v)})}/><Input p="Fiyat" type="number" v={job.price} c={(v:string)=>setJob({...job,price:Number(v)})}/><Input p="Kağıt Maliyeti" type="number" v={job.paper_cost} c={(v:string)=>setJob({...job,paper_cost:Number(v)})}/><Input p="Baskı Maliyeti" type="number" v={job.print_cost} c={(v:string)=>setJob({...job,print_cost:Number(v)})}/><Input p="Cilt Maliyeti" type="number" v={job.binding_cost} c={(v:string)=>setJob({...job,binding_cost:Number(v)})}/><Input p="Laminasyon Maliyeti" type="number" v={job.lamination_cost} c={(v:string)=>setJob({...job,lamination_cost:Number(v)})}/><Input p="İşçilik" type="number" v={job.labor_cost} c={(v:string)=>setJob({...job,labor_cost:Number(v)})}/><Input p="Teslim Tarihi" type="date" v={job.deadline} c={(v:string)=>setJob({...job,deadline:v})}/><Input p="Not" v={job.note} c={(v:string)=>setJob({...job,note:v})}/></div><div className="flex gap-3 mt-5"><button onClick={save} className="bg-blue-600 text-white px-5 py-3 rounded-lg font-bold">Kaydet</button><button onClick={()=>setJob(null)} className="bg-slate-200 px-5 py-3 rounded-lg font-bold">Kapat</button></div></div></div> }
function SimpleTable({head,rows}: any){ return <table className="w-full text-sm"><thead><tr className="border-b text-left text-slate-500">{head.map((h:string)=><th key={h} className="py-2">{h}</th>)}</tr></thead><tbody>{rows.map((r:any[],i:number)=><tr key={i} className="border-b">{r.map((x,k)=><td key={k} className="py-3">{x}</td>)}</tr>)}</tbody></table> }
