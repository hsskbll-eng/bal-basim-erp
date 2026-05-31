"use client"

import { useEffect, useMemo, useState } from "react"
import { supabase } from "@/lib/supabase"
import jsPDF from "jspdf"
import autoTable from "jspdf-autotable"
import * as XLSX from "xlsx"
import { openJobPrint } from "./pdfTemplates"

type Role = "admin" | "staff"
type Tab = "panel" | "newJob" | "customers" | "allJobs" | "invoice" | "reports" | "archive" | "stocks" | "deliveries" | "costs" | "costSettings"
type Status = "printing" | "cover" | "delivery" | "finished"
type InvoiceStatus = "waiting" | "invoiced" | "paid"
type Priority = "normal" | "urgent"

type Profile = { id: string; email: string; full_name: string; role: Role }
type Customer = { id: number; company: string; person: string | null; phone: string | null; email: string | null; address: string | null; note: string | null; created_at?: string }
type Log = { id?: number; job_id: number; text: string; created_at: string }
type Delivery = { id: number; job_id: number; amount: number; note: string | null; created_at: string }
type Stock = { id: number; name: string | null; type: string | null; quantity: number; unit: string | null; min_quantity: number; created_at?: string }

type CostSettings = {
  id: number
  paper_kg_price: number
  print_form_price: number
  folding_form_price: number
  cutting_price: number
  american_binding_price: number
  thread_binding_price: number
  staple_binding_price: number
  spiral_binding_price: number
  cover_print_price: number
  lamination_price: number
  waste_percent: number
  profit_percent: number
  vat_percent: number
}

type Job = {
  id: number
  customer_id: number | null
  customer_name: string | null
  job_name: string | null
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
  deadline: string | null
  note: string | null
  shrink_amount: string | null
  size: string | null
  page_count: string | null
  color: string | null
  print_type: string | null
  binding: string | null
  lamination: string | null
  cellophane: string | null
  inner_paper_gram: string | null
  inner_paper_type: string | null
  inner_paper_size: string | null
  inner_paper_amount: string | null
  cover_gram: string | null
  cover_type: string | null
  cover_paper_type: string | null
  cover_size: string | null
  cover_amount: string | null
  cover_inside_print: string | null
  created_at: string
  logs?: Log[]
  deliveries?: Delivery[]
}

const statusTitle: Record<Status, string> = { printing: "Baskıda", cover: "Kapak Takma", delivery: "Teslimat", finished: "Biten İşler" }
const invoiceTitle: Record<InvoiceStatus, string> = { waiting: "Fatura Bekliyor", invoiced: "Fatura Kesildi", paid: "Ödendi" }

const emptyCustomer = { company: "", person: "", phone: "", email: "", address: "", note: "" }
const emptyStock = { name: "", type: "Kağıt", quantity: "", unit: "Adet", min_quantity: "" }
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

const defaultCostSettings: CostSettings = {
  id: 0,
  paper_kg_price: 0,
  print_form_price: 0,
  folding_form_price: 0,
  cutting_price: 0,
  american_binding_price: 0,
  thread_binding_price: 0,
  staple_binding_price: 0,
  spiral_binding_price: 0,
  cover_print_price: 0,
  lamination_price: 0,
  waste_percent: 5,
  profit_percent: 30,
  vat_percent: 20,
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
  const [costSettings, setCostSettings] = useState<CostSettings | null>(null)

  const [customerForm, setCustomerForm] = useState(emptyCustomer)
  const [stockForm, setStockForm] = useState(emptyStock)
  const [jobForm, setJobForm] = useState(emptyJob)
  const [editJob, setEditJob] = useState<Job | null>(null)
  const [deliveryJob, setDeliveryJob] = useState<Job | null>(null)
  const [deliveryForm, setDeliveryForm] = useState({ amount: "", note: "" })

  const isAdmin = profile?.role === "admin"
  const activeJobs = jobs.filter((j) => !j.archived)
  const archivedJobs = jobs.filter((j) => j.archived)

  const stats = useMemo(() => ({
    total: activeJobs.length,
    urgent: activeJobs.filter((j) => j.priority === "urgent").length,
    late: activeJobs.filter((j) => isLate(j)).length,
    printing: activeJobs.filter((j) => j.status === "printing").length,
    cover: activeJobs.filter((j) => j.status === "cover").length,
    delivery: activeJobs.filter((j) => j.status === "delivery").length,
    finished: activeJobs.filter((j) => j.status === "finished").length,
    profit: activeJobs.reduce((s, j) => s + Number(j.profit || 0), 0),
  }), [jobs])

  const today = new Date().toISOString().slice(0, 10)
  const alerts = activeJobs
    .filter((j) => j.priority === "urgent" || j.deadline === today || (!!j.deadline && j.deadline < today && j.status !== "finished"))
    .map((j) => {
      const late = !!j.deadline && j.deadline < today && j.status !== "finished"
      const todayDelivery = j.deadline === today
      return {
        id: j.id,
        title: `${jobNo(j)} - ${j.customer_name || "Müşteri yok"}`,
        text: late ? "Teslim tarihi geçmiş" : todayDelivery ? "Bugün teslim edilecek" : "Acil iş",
        color: late ? "bg-red-100 text-red-700" : todayDelivery ? "bg-yellow-100 text-yellow-700" : "bg-orange-100 text-orange-700",
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
      { key: "costSettings" as Tab, text: "Maliyet Ayarları" },
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
    if (existing) {
      setProfile(existing as Profile)
      return
    }

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

    const { data: co, error: costError } = await supabase.from("cost_settings").select("*").order("id", { ascending: true }).limit(1)

    if (!costError && co && co.length > 0) {
      setCostSettings(co[0] as CostSettings)
    } else {
      const { data: inserted } = await supabase.from("cost_settings").insert({
        paper_kg_price: 0,
        print_form_price: 0,
        folding_form_price: 0,
        cutting_price: 0,
        american_binding_price: 0,
        thread_binding_price: 0,
        staple_binding_price: 0,
        spiral_binding_price: 0,
        cover_print_price: 0,
        lamination_price: 0,
        waste_percent: 5,
        profit_percent: 30,
        vat_percent: 20,
      }).select().single()
      setCostSettings((inserted as CostSettings) || defaultCostSettings)
    }

    const jobsWith = ((js || []) as Job[]).map((j) => ({
      ...j,
      logs: (ls || []).filter((l: any) => l.job_id === j.id),
      deliveries: (ds || []).filter((d: any) => d.job_id === j.id),
    }))

    setCustomers((cs || []) as Customer[])
    setJobs(jobsWith)
    setStocks((st || []) as Stock[])
    if ((cs || [])[0]) setJobForm((p) => ({ ...p, customer_id: (cs || [])[0].id }))
  }

  async function login() {
    const { error } = await supabase.auth.signInWithPassword({ email: auth.email, password: auth.password })
    if (error) alert(error.message)
  }

  async function register() {
    const { error } = await supabase.auth.signUp({ email: auth.email, password: auth.password, options: { data: { full_name: auth.fullName } } })
    if (error) alert(error.message)
    else alert("Kayıt oluşturuldu. E-posta doğrulama kapalıysa direkt giriş yapabilirsin.")
  }

  async function logout() {
    await supabase.auth.signOut()
    setProfile(null)
    setUser(null)
  }

  async function addLog(jobId: number, text: string) {
    await supabase.from("logs").insert({ job_id: jobId, text })
  }

  async function addCustomer() {
    if (!customerForm.company) return alert("Firma adı zorunlu.")
    const { error } = await supabase.from("customers").insert(customerForm)
    if (error) return alert(error.message)
    setCustomerForm(emptyCustomer)
    await loadAll()
  }

  function calcTotals(data: typeof emptyJob) {
    const total_cost = [data.paper_cost, data.print_cost, data.binding_cost, data.lamination_cost, data.labor_cost].reduce((s, v) => s + Number(v || 0), 0)
    const profit = Number(data.price || 0) - total_cost
    return { total_cost, profit }
  }

  async function addJob() {
    const customer = customers.find((c) => c.id === Number(jobForm.customer_id))
    if (!customer || !jobForm.job_name) return alert("Müşteri ve iş adı zorunlu.")

    const totals = calcTotals(jobForm)
    const payload = {
      customer_id: customer.id,
      customer_name: customer.company,
      job_name: jobForm.job_name,
      quantity: Number(jobForm.quantity || 0),
      delivered: 0,
      price: Number(jobForm.price || 0),
      paper_cost: Number(jobForm.paper_cost || 0),
      print_cost: Number(jobForm.print_cost || 0),
      binding_cost: Number(jobForm.binding_cost || 0),
      lamination_cost: Number(jobForm.lamination_cost || 0),
      labor_cost: Number(jobForm.labor_cost || 0),
      total_cost: totals.total_cost,
      profit: totals.profit,
      status: "printing",
      invoice_status: "waiting",
      priority: jobForm.priority,
      archived: false,
      deadline: jobForm.deadline,
      note: jobForm.note,
      shrink_amount: jobForm.shrink_amount,
      size: jobForm.size,
      page_count: jobForm.page_count,
      color: jobForm.color,
      print_type: jobForm.print_type,
      binding: jobForm.binding,
      lamination: jobForm.lamination,
      cellophane: jobForm.cellophane,
      inner_paper_gram: jobForm.inner_paper_gram,
      inner_paper_type: jobForm.inner_paper_type,
      inner_paper_size: jobForm.inner_paper_size,
      inner_paper_amount: jobForm.inner_paper_amount,
      cover_gram: jobForm.cover_gram,
      cover_type: jobForm.cover_type,
      cover_paper_type: jobForm.cover_paper_type,
      cover_size: jobForm.cover_size,
      cover_amount: jobForm.cover_amount,
      cover_inside_print: jobForm.cover_inside_print,
    }

    const { data, error } = await supabase.from("jobs").insert(payload).select().single()
    if (error) return alert(error.message)
    await addLog(data.id, "İş oluşturuldu.")
    setJobForm({ ...emptyJob, customer_id: customers[0]?.id || 0 })
    setTab("panel")
    await loadAll()
  }

  async function saveEditJob() {
    if (!editJob) return
    const total_cost = Number(editJob.paper_cost || 0) + Number(editJob.print_cost || 0) + Number(editJob.binding_cost || 0) + Number(editJob.lamination_cost || 0) + Number(editJob.labor_cost || 0)
    const profit = Number(editJob.price || 0) - total_cost
    const { logs, deliveries, ...clean } = editJob
    const { error } = await supabase.from("jobs").update({ ...clean, total_cost, profit }).eq("id", editJob.id)
    if (error) return alert(error.message)
    await addLog(editJob.id, "İş bilgileri düzenlendi.")
    setEditJob(null)
    await loadAll()
  }

  async function updateJob(job: Job, updates: Partial<Job>, logText: string) {
    const { error } = await supabase.from("jobs").update(updates).eq("id", job.id)
    if (error) return alert(error.message)
    await addLog(job.id, logText)
    await loadAll()
  }

  async function nextJob(job: Job) {
    if (job.status === "printing") return updateJob(job, { status: "cover" }, "Baskı tamamlandı, kapak takmaya geçti.")
    if (job.status === "cover") return updateJob(job, { status: "delivery" }, "Kapak takıldı, teslimata geçti.")
    if (job.status === "delivery") {
      setDeliveryJob(job)
      setDeliveryForm({ amount: "", note: "" })
    }
  }

  async function prevJob(job: Job) {
    if (job.status === "cover") return updateJob(job, { status: "printing" }, "Geri alındı: Baskıya döndü.")
    if (job.status === "delivery") return updateJob(job, { status: "cover" }, "Geri alındı: Kapak takmaya döndü.")
    if (job.status === "finished") return updateJob(job, { status: "delivery" }, "Geri alındı: Teslimata döndü.")
  }

  async function archiveJob(job: Job) {
    if (!confirm("Arşive taşınsın mı?")) return
    await updateJob(job, { archived: true }, "İş arşive taşındı.")
  }

  async function restoreJob(job: Job) {
    await updateJob(job, { archived: false }, "İş arşivden geri alındı.")
  }

  async function deleteArchivedJob(job: Job) {
    if (!confirm(`${jobNo(job)} numaralı işi tamamen silmek istiyor musun?`)) return
    await supabase.from("deliveries").delete().eq("job_id", job.id)
    await supabase.from("logs").delete().eq("job_id", job.id)
    await supabase.from("jobs").delete().eq("id", job.id)
    await loadAll()
  }

  async function updateInvoice(job: Job, invoice_status: InvoiceStatus) {
    await updateJob(job, { invoice_status }, `Fatura durumu: ${invoiceTitle[invoice_status]}`)
  }

  async function copyJob(job: Job) {
    const { id, logs, deliveries, created_at, ...copy } = job
    const { data, error } = await supabase.from("jobs").insert({ ...copy, job_name: `${job.job_name} - Kopya`, delivered: 0, status: "printing", invoice_status: "waiting", archived: false }).select().single()
    if (error) return alert(error.message)
    if (data) await addLog(data.id, `${jobNo(job)} üzerinden kopyalandı.`)
    await loadAll()
  }

  async function addStock() {
    if (!stockForm.name) return alert("Stok adı zorunlu.")
    const { error } = await supabase.from("stocks").insert({ name: stockForm.name, type: stockForm.type, quantity: Number(stockForm.quantity || 0), unit: stockForm.unit, min_quantity: Number(stockForm.min_quantity || 0) })
    if (error) return alert(error.message)
    setStockForm(emptyStock)
    await loadAll()
  }

  async function moveStock(stock: Stock, movement_type: "Giriş" | "Çıkış") {
    const amount = Number(prompt(`${movement_type} miktarı`) || 0)
    if (amount <= 0) return
    const newQty = movement_type === "Giriş" ? Number(stock.quantity) + amount : Number(stock.quantity) - amount
    await supabase.from("stocks").update({ quantity: newQty }).eq("id", stock.id)
    await loadAll()
  }

  async function saveCostSettings() {
    if (!costSettings) return
    if (costSettings.id === 0) {
      const { data, error } = await supabase.from("cost_settings").insert({ ...costSettings, id: undefined }).select().single()
      if (error) return alert(error.message)
      setCostSettings(data as CostSettings)
    } else {
      const { error } = await supabase.from("cost_settings").update(costSettings).eq("id", costSettings.id)
      if (error) return alert(error.message)
    }
    alert("Maliyet ayarları kaydedildi.")
    await loadAll()
  }

  function calculateJobCosts() {
    if (!costSettings) return alert("Maliyet ayarları yüklenmedi.")
    const quantity = Number(jobForm.quantity || 0)
    const pages = Number(jobForm.page_count || 0)
    if (quantity <= 0 || pages <= 0) return alert("Adet ve sayfa sayısı gir.")

    const forms = Math.max(Math.ceil(pages / 16), 1)
    const paperBase = quantity * forms * Number(costSettings.paper_kg_price || 0)
    const waste = paperBase * (Number(costSettings.waste_percent || 0) / 100)
    const paper_cost = Math.round(paperBase + waste)
    const print_cost = Math.round(forms * Number(costSettings.print_form_price || 0))
    const folding_cost = Math.round(forms * Number(costSettings.folding_form_price || 0))

    const bindingText = String(jobForm.binding || "").toLowerCase()
    let bindingUnit = 0
    if (bindingText.includes("amerikan")) bindingUnit = Number(costSettings.american_binding_price || 0)
    else if (bindingText.includes("iplik")) bindingUnit = Number(costSettings.thread_binding_price || 0)
    else if (bindingText.includes("tel")) bindingUnit = Number(costSettings.staple_binding_price || 0)
    else if (bindingText.includes("spiral")) bindingUnit = Number(costSettings.spiral_binding_price || 0)

    const binding_cost = Math.round(quantity * bindingUnit)
    const lamination_cost = jobForm.lamination === "Var" ? Math.round(quantity * Number(costSettings.lamination_price || 0)) : 0
    const other = Number(costSettings.cutting_price || 0) + Number(costSettings.cover_print_price || 0)
    const total_cost = paper_cost + print_cost + folding_cost + binding_cost + lamination_cost + other
    const suggested_price = Math.round(total_cost * (1 + Number(costSettings.profit_percent || 0) / 100))

    setJobForm({
      ...jobForm,
      paper_cost: String(paper_cost),
      print_cost: String(print_cost),
      binding_cost: String(binding_cost + folding_cost),
      lamination_cost: String(lamination_cost),
      labor_cost: String(other),
      price: String(suggested_price),
    })

    alert(`Forma: ${forms}\nToplam maliyet: ${total_cost.toLocaleString("tr-TR")} ₺\nÖnerilen satış: ${suggested_price.toLocaleString("tr-TR")} ₺`)
  }

  function makeJobPdf(job: Job) { openJobPrint(job) }

  function makeDeliveryPdf(job: Job) {
    const w = window.open("", "_blank")
    if (!w) return alert("Açılır pencere engellendi.")
    const remaining = Math.max(Number(job.quantity || 0) - Number(job.delivered || 0), 0)
    w.document.write(`<html lang="tr"><head><meta charset="UTF-8"/><title>Teslim Fişi ${jobNo(job)}</title><style>body{margin:0;background:#d1d5db;font-family:Arial,Helvetica,sans-serif;color:#071d35}.page{width:210mm;min-height:297mm;background:white;margin:20px auto;padding:14mm}.logo{text-align:center;margin-bottom:10px}.logo img{width:220px}.title{background:#071d35;color:white;text-align:center;padding:14px;font-size:28px;font-weight:900;margin-bottom:20px}table{width:100%;border-collapse:collapse}td{border:1px solid #9ca3af;padding:12px;font-size:14px}.dark{background:#071d35;color:white;font-weight:900;width:220px}.big{font-size:24px;font-weight:900;text-align:center}.notes{height:140px;vertical-align:top}.signs{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:60px}.sign{text-align:center;font-weight:900}.line{border-bottom:2px solid black;margin-top:60px}@media print{body{background:white}.page{margin:0}}</style></head><body><div class="page"><div class="logo"><img src="/logo.png" /></div><div class="title">TESLİM FİŞİ</div><table><tr><td class="dark">Teslim Fiş No</td><td class="big">${jobNo(job)}</td></tr><tr><td class="dark">Firma</td><td>${job.customer_name || "-"}</td></tr><tr><td class="dark">İş Adı</td><td>${job.job_name || "-"}</td></tr><tr><td class="dark">Toplam Adet</td><td>${job.quantity}</td></tr><tr><td class="dark">Teslim Edilen</td><td>${job.delivered}</td></tr><tr><td class="dark">Kalan</td><td>${remaining}</td></tr><tr><td class="dark">Teslim Tarihi</td><td>${new Date().toLocaleDateString("tr-TR")}</td></tr><tr><td class="dark">Teslim Notu</td><td class="notes">${job.note || ""}</td></tr></table><div class="signs"><div class="sign">Teslim Eden<div class="line"></div></div><div class="sign">Teslim Alan<div class="line"></div></div></div></div><script>setTimeout(() => window.print(), 500)</script></body></html>`)
    w.document.close()
  }

  function makeMonthlyReportPdf(monthJobs: Job[]) {
    const pdf = new jsPDF()
    const total = monthJobs.reduce((s, j) => s + Number(j.price || 0), 0)
    const cost = monthJobs.reduce((s, j) => s + Number(j.total_cost || 0), 0)
    const profit = total - cost
    pdf.setFillColor(7, 29, 53); pdf.rect(0, 0, 210, 30, "F"); pdf.setTextColor(255, 255, 255); pdf.setFontSize(18); pdf.text("BAL BASIM YAYIN", 14, 18); pdf.setTextColor(0, 0, 0)
    autoTable(pdf, { startY: 40, head: [["Özet", "Bilgi"]], body: [["Ay", monthFilter], ["Toplam İş", monthJobs.length], ["Toplam Ciro", `${total.toLocaleString("tr-TR")} TL`], ["Toplam Maliyet", `${cost.toLocaleString("tr-TR")} TL`], ["Toplam Kâr", `${profit.toLocaleString("tr-TR")} TL`]], headStyles: { fillColor: [7, 29, 53], textColor: 255 } })
    autoTable(pdf, { startY: 100, head: [["İş No", "Müşteri", "İş", "Adet", "Ciro", "Maliyet", "Kâr"]], body: monthJobs.map((j) => [jobNo(j), j.customer_name, j.job_name, j.quantity, `${j.price} TL`, `${j.total_cost} TL`, `${j.profit} TL`]), headStyles: { fillColor: [7, 29, 53], textColor: 255 } })
    const url = URL.createObjectURL(pdf.output("blob")); window.open(url, "_blank")
  }

  function exportJobsExcel() {
    const rows = activeJobs.map((j) => ({ İşNo: jobNo(j), Müşteri: j.customer_name, İş: j.job_name, Adet: j.quantity, Teslim: j.delivered, Durum: statusTitle[j.status], Fatura: invoiceTitle[j.invoice_status], Fiyat: j.price, Maliyet: j.total_cost, Kar: j.profit }))
    const wb = XLSX.utils.book_new(); const ws = XLSX.utils.json_to_sheet(rows); XLSX.utils.book_append_sheet(wb, ws, "İşler"); XLSX.writeFile(wb, "bal-basim-isler.xlsx")
  }

  function isLate(job: Job) { if (!job.deadline || job.status === "finished") return false; return new Date(job.deadline) < new Date(new Date().toDateString()) }
  const searchedJobs = activeJobs.filter((j) => `${j.id} ${jobNo(j)} ${j.customer_name || ""} ${j.job_name || ""} ${statusTitle[j.status]}`.toLowerCase().includes(search.toLowerCase()))
  const monthlyJobs = activeJobs.filter((j) => j.created_at?.slice(0, 7) === monthFilter)

  if (loading) return <div className="p-10 text-2xl font-black">Yükleniyor...</div>
  if (!user) return <Login auth={auth} setAuth={setAuth} authMode={authMode} setAuthMode={setAuthMode} login={login} register={register} />

  return (
    <main className="min-h-screen bg-[#f4f7fb] text-slate-900">
      <datalist id="paper-sizes"><option value="57x82" /><option value="64x90" /><option value="68x100" /><option value="70x100" /><option value="Özel Ölçü" /></datalist>
      <datalist id="paper-types"><option value="1. Hamur" /><option value="2. Hamur" /><option value="Kuşe" /><option value="Bristol" /><option value="Kitap Kağıdı" /></datalist>
      <datalist id="bindings"><option value="Amerikan Cilt" /><option value="İplik Dikiş" /><option value="Tel Dikiş" /><option value="Spiral" /><option value="Sert Kapak" /></datalist>
      <datalist id="grams"><option value="55 gr" /><option value="60 gr" /><option value="70 gr" /><option value="80 gr" /><option value="90 gr" /><option value="105 gr" /><option value="115 gr" /><option value="130 gr" /><option value="170 gr" /><option value="250 gr" /><option value="300 gr" /></datalist>
      <datalist id="print-types"><option value="Tek Renk" /><option value="2 Renk" /><option value="4+0" /><option value="4+1" /><option value="4+4" /></datalist>

      <aside className="hidden md:block fixed left-0 top-0 h-screen overflow-y-auto w-[270px] bg-[#071d35] text-white p-5">
        <div className="mb-8"><img src="/logo.png" className="w-44 mb-3 bg-white rounded-lg p-2" /><div className="text-sm text-slate-300">Premium ERP</div></div>
        <div className="mb-5 text-sm"><b>{profile?.full_name}</b><br /><span className="text-slate-300">{profile?.role === "admin" ? "Yönetici" : "Personel"}</span></div>
        {visibleTabs.map((i) => <button key={i.key} onClick={() => setTab(i.key)} className={`w-full text-left px-4 py-3 rounded-lg mb-2 text-sm ${tab === i.key ? "bg-blue-600" : "text-slate-200 hover:bg-slate-800"}`}>{i.text}</button>)}
        <button onClick={logout} className="w-full mt-6 bg-red-600 py-3 rounded-lg font-bold">Çıkış</button>
      </aside>

      <MobileHeader profile={profile} logout={logout} />

      <section className="md:ml-[270px] p-4 md:p-7 pb-28 md:pb-7">
        <h1 className="text-2xl font-black mb-6">{tabTitle(tab)}</h1>
        {tab === "panel" && <><MobilePanelHero stats={stats} alerts={alerts} isAdmin={isAdmin} /><Alerts alerts={alerts} /><div className="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-8 gap-3 md:gap-4 mb-7"><Stat title="Toplam" value={stats.total} /><Stat title="Acil" value={stats.urgent} /><Stat title="Geciken" value={stats.late} /><Stat title="Baskı" value={stats.printing} /><Stat title="Kapak" value={stats.cover} /><Stat title="Teslim" value={stats.delivery} /><Stat title="Biten" value={stats.finished} />{isAdmin && <Stat title="Kâr" valueText={`${stats.profit.toLocaleString("tr-TR")} ₺`} />}</div><div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4 md:gap-5">{(["printing", "cover", "delivery", "finished"] as Status[]).map((s) => <Column key={s} status={s} jobs={activeJobs.filter((j) => j.status === s)} isAdmin={isAdmin} nextJob={nextJob} prevJob={prevJob} archiveJob={archiveJob} copyJob={copyJob} setEditJob={setEditJob} makeJobPdf={makeJobPdf} makeDeliveryPdf={makeDeliveryPdf} />)}</div></>}
        {tab === "newJob" && isAdmin && <JobForm customers={customers} jobForm={jobForm} setJobForm={setJobForm} addJob={addJob} calculateJobCosts={calculateJobCosts} />}
        {tab === "customers" && isAdmin && <Customers customers={customers} customerForm={customerForm} setCustomerForm={setCustomerForm} addCustomer={addCustomer} jobs={jobs} loadAll={loadAll} />}
        {tab === "allJobs" && <Panel><div className="flex justify-between mb-4"><input className="border rounded-lg p-2 w-full md:w-[420px]" placeholder="Ara..." value={search} onChange={(e) => setSearch(e.target.value)} />{isAdmin && <button onClick={exportJobsExcel} className="bg-green-600 text-white px-4 py-2 rounded-lg font-bold">Excel Aktar</button>}</div><JobsTable jobs={searchedJobs} isAdmin={isAdmin} makeJobPdf={makeJobPdf} /></Panel>}
        {tab === "invoice" && isAdmin && <Panel><InvoiceTable jobs={activeJobs.filter((j) => j.status === "finished")} updateInvoice={updateInvoice} makeJobPdf={makeJobPdf} /></Panel>}
        {tab === "costs" && isAdmin && <Panel><JobsTable jobs={activeJobs} isAdmin={true} makeJobPdf={makeJobPdf} showCosts /></Panel>}
        {tab === "costSettings" && isAdmin && <CostSettingsPanel settings={costSettings} setSettings={setCostSettings} save={saveCostSettings} />}
        {tab === "stocks" && isAdmin && <Stocks stocks={stocks} stockForm={stockForm} setStockForm={setStockForm} addStock={addStock} moveStock={moveStock} />}
        {tab === "deliveries" && isAdmin && <Panel><Deliveries jobs={activeJobs} /></Panel>}
        {tab === "reports" && isAdmin && <Panel><div className="flex justify-between mb-5"><input type="month" className="border rounded-lg p-2" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)} /><button onClick={() => makeMonthlyReportPdf(monthlyJobs)} className="bg-slate-900 text-white px-4 py-2 rounded-lg font-bold">Aylık PDF</button></div><JobsTable jobs={monthlyJobs} isAdmin={true} makeJobPdf={makeJobPdf} showCosts /></Panel>}
        {tab === "archive" && isAdmin && <Panel>{archivedJobs.map((j) => <div key={j.id} className="border rounded-lg p-3 mb-2 flex justify-between items-center"><b>{jobNo(j)} {j.customer_name} - {j.job_name}</b><div className="flex gap-2"><button onClick={() => restoreJob(j)} className="bg-blue-600 text-white px-3 py-2 rounded">Geri Al</button><button onClick={() => deleteArchivedJob(j)} className="bg-red-600 text-white px-3 py-2 rounded">Tamamen Sil</button></div></div>)}</Panel>}
      </section>

      <MobileNav tab={tab} setTab={setTab} isAdmin={isAdmin} />

      {editJob && <EditModal job={editJob} setJob={setEditJob} save={saveEditJob} />}
      {deliveryJob && <DeliveryModal job={deliveryJob} form={deliveryForm} setForm={setDeliveryForm} close={() => setDeliveryJob(null)} save={async () => { const amount = Number(deliveryForm.amount || 0); if (amount <= 0) return alert("Teslim adedi gir."); const delivered = Number(deliveryJob.delivered || 0) + amount; await supabase.from("deliveries").insert({ job_id: deliveryJob.id, amount, note: deliveryForm.note }); await updateJob(deliveryJob, { delivered, status: delivered >= deliveryJob.quantity ? "finished" : "delivery" }, `${amount} adet teslim edildi.`); setDeliveryJob(null) }} />}
    </main>
  )
}

function MobileHeader({ profile, logout }: any) {
  return (
    <div className="md:hidden sticky top-0 z-40 bg-[#071d35] text-white shadow-xl">
      <div className="px-4 pt-4 pb-5">
        <div className="flex items-center justify-between gap-3">
          <div className="flex items-center gap-3 min-w-0">
            <img src="/logo.png" className="w-24 bg-white rounded-2xl p-2 shadow" />
            <div className="min-w-0">
              <div className="text-xs text-slate-300">BAL BASIM ERP</div>
              <div className="font-black text-base truncate">{profile?.full_name}</div>
              <div className="text-xs text-slate-300">{profile?.role === "admin" ? "Yönetici" : "Personel"}</div>
            </div>
          </div>
          <button onClick={logout} className="bg-red-600 px-4 py-2 rounded-2xl text-xs font-black shrink-0 shadow">Çıkış</button>
        </div>
      </div>
    </div>
  )
}

function MobileNav({ tab, setTab, isAdmin }: any) {
  const items = [
    { key: "panel", text: "Panel", icon: "📋" },
    ...(isAdmin ? [{ key: "newJob", text: "Yeni", icon: "➕" }] : []),
    { key: "allJobs", text: "İşler", icon: "📚" },
    ...(isAdmin ? [{ key: "stocks", text: "Stok", icon: "📦" }, { key: "customers", text: "Müşteri", icon: "👥" }] : []),
  ]

  return (
    <div className="md:hidden fixed bottom-0 left-0 right-0 z-50 bg-white/95 backdrop-blur border-t shadow-2xl px-2 pt-2 pb-3">
      <div className="grid grid-cols-5 gap-1">
        {items.slice(0, 5).map((i: any) => (
          <button
            key={i.key}
            onClick={() => setTab(i.key)}
            className={`py-2 rounded-2xl text-[11px] font-black flex flex-col items-center gap-1 ${tab === i.key ? "bg-blue-600 text-white shadow-lg" : "bg-slate-100 text-slate-700"}`}
          >
            <span className="text-lg leading-none">{i.icon}</span>
            <span>{i.text}</span>
          </button>
        ))}
      </div>
    </div>
  )
}

function MobilePanelHero({ stats, alerts, isAdmin }: any) {
  return (
    <div className="md:hidden -mx-4 -mt-4 mb-5 bg-[#071d35] text-white px-4 pb-6 rounded-b-[32px] shadow-xl">
      <div className="pt-2">
        <div className="text-sm text-slate-300">Bugünkü ERP Durumu</div>
        <div className="text-3xl font-black mt-1">İş Takip Paneli</div>
      </div>
      <div className="grid grid-cols-3 gap-3 mt-5">
        <div className="bg-white text-slate-900 rounded-2xl p-4 shadow"><div className="text-xs font-bold text-slate-500">Toplam</div><div className="text-3xl font-black">{stats.total}</div></div>
        <div className="bg-white text-slate-900 rounded-2xl p-4 shadow"><div className="text-xs font-bold text-slate-500">Acil</div><div className="text-3xl font-black text-red-600">{stats.urgent}</div></div>
        <div className="bg-white text-slate-900 rounded-2xl p-4 shadow"><div className="text-xs font-bold text-slate-500">Teslim</div><div className="text-3xl font-black">{stats.delivery}</div></div>
      </div>
      {alerts.length > 0 && (
        <div className="mt-4 bg-white/10 border border-white/10 rounded-2xl p-3">
          <div className="text-xs text-slate-300 font-bold mb-1">Canlı Bildirim</div>
          <div className="text-sm font-black truncate">{alerts[0].title}</div>
          <div className="text-xs text-slate-300">{alerts[0].text}</div>
        </div>
      )}
      {isAdmin && <div className="mt-4 text-sm text-green-300 font-black">Kâr: {stats.profit.toLocaleString("tr-TR")} ₺</div>}
    </div>
  )
}

function jobNo(job: Job) { const year = String(new Date(job.created_at).getFullYear()).slice(2); return `${String(job.id).padStart(4, "0")}-${year}` }
function Login({ auth, setAuth, authMode, setAuthMode, login, register }: any) { return <main className="min-h-screen bg-[#071d35] flex items-center justify-center"><div className="bg-white rounded-2xl p-8 w-[420px]"><img src="/logo.png" className="w-56 mx-auto mb-4" /><p className="text-slate-500 mb-6 text-center">Premium ERP Giriş</p>{authMode === "register" && <Input p="Ad Soyad" v={auth.fullName} c={(v: string) => setAuth({ ...auth, fullName: v })} />}<div className="mt-3"><Input p="E-posta" v={auth.email} c={(v: string) => setAuth({ ...auth, email: v })} /></div><div className="mt-3"><Input p="Şifre" type="password" v={auth.password} c={(v: string) => setAuth({ ...auth, password: v })} onKeyDown={(e: any) => { if (e.key === "Enter") login() }} /></div><button onClick={authMode === "login" ? login : register} className="w-full mt-5 bg-blue-600 text-white py-3 rounded-lg font-bold">{authMode === "login" ? "Giriş Yap" : "Kayıt Ol"}</button><button onClick={() => setAuthMode(authMode === "login" ? "register" : "login")} className="w-full mt-3 text-blue-600">{authMode === "login" ? "Hesap oluştur" : "Girişe dön"}</button></div></main> }
function tabTitle(tab: Tab) { return { panel: "📋 İş Takip Paneli", newJob: "➕ Yeni İş Girişi", customers: "👥 Müşteriler", allJobs: "📚 Tüm İşler", invoice: "💰 Faturalar", reports: "📊 Raporlar", archive: "🗄 Arşiv", stocks: "📦 Stok", deliveries: "🚚 Teslimat Geçmişi", costs: "📈 Maliyet / Kâr", costSettings: "⚙️ Maliyet Ayarları" }[tab] }
function Panel({ children }: any) { return <div className="bg-white border rounded-xl p-4 md:p-5 overflow-x-auto">{children}</div> }
function Section({ title, children }: any) { return <Panel><h2 className="font-black text-lg md:text-xl mb-4">{title}</h2><div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-3">{children}</div></Panel> }
function Input({ p, v, c, type = "text", onKeyDown, list }: any) { return <input className="border rounded-lg p-2 text-sm w-full" placeholder={p} value={v || ""} type={type} list={list} onFocus={(e) => e.target.select()} onChange={(e) => c(e.target.value)} onKeyDown={onKeyDown} /> }
function RadioGroup({ label, value, options, onChange }: any) { return <div><label className="text-xs font-bold text-slate-600">{label}</label><div className="flex gap-2 mt-1">{options.map((o: any) => <button key={o.value} type="button" onClick={() => onChange(o.value)} className={`px-4 py-2 rounded-lg border text-sm font-bold ${value === o.value ? "bg-blue-600 text-white" : "bg-white"}`}>{o.label}</button>)}</div></div> }
function Stat({ title, value, valueText }: any) { return <div className="bg-white border rounded-xl p-5"><div className="text-sm text-slate-500 font-bold">{title}</div><div className="text-2xl font-black mt-2">{valueText || value}</div></div> }
function Alerts({ alerts }: any) { if (!alerts.length) return null; return <div className="bg-white border rounded-xl p-5 mb-6"><h2 className="font-black text-lg mb-3">🔔 Canlı Bildirimler</h2><div className="grid grid-cols-3 gap-3">{alerts.map((a: any) => <div key={a.id} className={`rounded-lg p-3 font-bold ${a.color}`}><div>{a.title}</div><div className="text-sm opacity-80">{a.text}</div></div>)}</div></div> }
function JobForm({ customers, jobForm, setJobForm, addJob, calculateJobCosts }: any) { return <div className="space-y-5"><Section title="Genel İş Bilgileri"><select className="border rounded-lg p-2 text-sm" value={jobForm.customer_id} onChange={(e) => setJobForm({ ...jobForm, customer_id: Number(e.target.value) })}>{customers.map((c: Customer) => <option key={c.id} value={c.id}>{c.company}</option>)}</select><Input p="İş Adı" v={jobForm.job_name} c={(v: string) => setJobForm({ ...jobForm, job_name: v })} /><Input p="Adet" type="number" v={jobForm.quantity} c={(v: string) => setJobForm({ ...jobForm, quantity: v })} /><Input p="Fiyat" type="number" v={jobForm.price} c={(v: string) => setJobForm({ ...jobForm, price: v })} /><Input p="Teslim Tarihi" type="date" v={jobForm.deadline} c={(v: string) => setJobForm({ ...jobForm, deadline: v })} /><RadioGroup label="Öncelik" value={jobForm.priority} options={[{ label: "Normal", value: "normal" }, { label: "Acil", value: "urgent" }]} onChange={(v: string) => setJobForm({ ...jobForm, priority: v })} /></Section><Section title="Maliyet Bilgileri"><Input p="Kağıt Maliyeti" type="number" v={jobForm.paper_cost} c={(v: string) => setJobForm({ ...jobForm, paper_cost: v })} /><Input p="Baskı Maliyeti" type="number" v={jobForm.print_cost} c={(v: string) => setJobForm({ ...jobForm, print_cost: v })} /><Input p="Cilt Maliyeti" type="number" v={jobForm.binding_cost} c={(v: string) => setJobForm({ ...jobForm, binding_cost: v })} /><Input p="Laminasyon Maliyeti" type="number" v={jobForm.lamination_cost} c={(v: string) => setJobForm({ ...jobForm, lamination_cost: v })} /><Input p="İşçilik" type="number" v={jobForm.labor_cost} c={(v: string) => setJobForm({ ...jobForm, labor_cost: v })} /></Section><Section title="Baskı Bilgileri"><Input p="Shrink Adedi" v={jobForm.shrink_amount} c={(v: string) => setJobForm({ ...jobForm, shrink_amount: v })} /><Input p="Ebat" v={jobForm.size} list="paper-sizes" c={(v: string) => setJobForm({ ...jobForm, size: v })} /><Input p="Sayfa Sayısı" v={jobForm.page_count} c={(v: string) => setJobForm({ ...jobForm, page_count: v })} /><Input p="Renk" v={jobForm.color} list="print-types" c={(v: string) => setJobForm({ ...jobForm, color: v })} /><Input p="Baskı Tipi" v={jobForm.print_type} list="print-types" c={(v: string) => setJobForm({ ...jobForm, print_type: v })} /><Input p="Cilt Şekli" v={jobForm.binding} list="bindings" c={(v: string) => setJobForm({ ...jobForm, binding: v })} /><RadioGroup label="Laminasyon" value={jobForm.lamination} options={[{ label: "Yok", value: "Yok" }, { label: "Var", value: "Var" }]} onChange={(v: string) => setJobForm({ ...jobForm, lamination: v })} /><RadioGroup label="Selefon" value={jobForm.cellophane} options={[{ label: "Yok", value: "Yok" }, { label: "Mat", value: "Mat" }, { label: "Parlak", value: "Parlak" }]} onChange={(v: string) => setJobForm({ ...jobForm, cellophane: v })} /></Section><Section title="İç Kağıt Bilgileri"><Input p="Gramaj" v={jobForm.inner_paper_gram} list="grams" c={(v: string) => setJobForm({ ...jobForm, inner_paper_gram: v })} /><Input p="Türü" v={jobForm.inner_paper_type} list="paper-types" c={(v: string) => setJobForm({ ...jobForm, inner_paper_type: v })} /><Input p="Ebatı" v={jobForm.inner_paper_size} list="paper-sizes" c={(v: string) => setJobForm({ ...jobForm, inner_paper_size: v })} /><Input p="Miktarı" v={jobForm.inner_paper_amount} c={(v: string) => setJobForm({ ...jobForm, inner_paper_amount: v })} /></Section><Section title="Kapak Kağıdı Bilgileri"><Input p="Gramaj" v={jobForm.cover_gram} list="grams" c={(v: string) => setJobForm({ ...jobForm, cover_gram: v })} /><Input p="Türü" v={jobForm.cover_type} list="paper-types" c={(v: string) => setJobForm({ ...jobForm, cover_type: v })} /><Input p="Kağıt Türü" v={jobForm.cover_paper_type} list="paper-types" c={(v: string) => setJobForm({ ...jobForm, cover_paper_type: v })} /><Input p="Ebat" v={jobForm.cover_size} list="paper-sizes" c={(v: string) => setJobForm({ ...jobForm, cover_size: v })} /><Input p="Miktar" v={jobForm.cover_amount} c={(v: string) => setJobForm({ ...jobForm, cover_amount: v })} /><RadioGroup label="Kapak İçi Baskı" value={jobForm.cover_inside_print} options={[{ label: "Yok", value: "Yok" }, { label: "Var", value: "Var" }]} onChange={(v: string) => setJobForm({ ...jobForm, cover_inside_print: v })} /></Section><textarea className="border rounded-lg p-3 w-full" placeholder="Not" value={jobForm.note} onChange={(e) => setJobForm({ ...jobForm, note: e.target.value })} /><div className="flex gap-3"><button onClick={calculateJobCosts} className="bg-slate-900 text-white px-6 py-3 rounded-lg font-bold">Maliyeti Hesapla</button><button onClick={addJob} className="bg-blue-600 text-white px-6 py-3 rounded-lg font-bold">İşi Kaydet</button></div></div> }
function Customers({ customers, customerForm, setCustomerForm, addCustomer, jobs, loadAll }: any) { return <div className="grid grid-cols-1 md:grid-cols-3 gap-5"><Panel><h2 className="font-black text-xl mb-4">Müşteri Kayıt</h2><div className="space-y-3"><Input p="Firma" v={customerForm.company} c={(v: string) => setCustomerForm({ ...customerForm, company: v })} /><Input p="Yetkili" v={customerForm.person} c={(v: string) => setCustomerForm({ ...customerForm, person: v })} /><Input p="Telefon" v={customerForm.phone} c={(v: string) => setCustomerForm({ ...customerForm, phone: v })} /><Input p="E-posta" v={customerForm.email} c={(v: string) => setCustomerForm({ ...customerForm, email: v })} /><Input p="Adres" v={customerForm.address} c={(v: string) => setCustomerForm({ ...customerForm, address: v })} /><Input p="Not" v={customerForm.note} c={(v: string) => setCustomerForm({ ...customerForm, note: v })} /><button onClick={addCustomer} className="bg-blue-600 text-white px-5 py-3 rounded-lg font-bold">Müşteri Ekle</button></div></Panel><div className="md:col-span-2"><Panel><table className="w-full text-sm min-w-[700px]"><thead><tr className="border-b text-left text-slate-500"><th className="py-2">Firma</th><th>Yetkili</th><th>Telefon</th><th>Adres</th><th>İşlem</th></tr></thead><tbody>{customers.map((c: Customer) => <tr key={c.id} className="border-b"><td className="py-3 font-bold">{c.company}</td><td>{c.person}</td><td>{c.phone}</td><td>{c.address}</td><td><button onClick={async () => { const hasJobs = jobs.some((j: Job) => j.customer_id === c.id); if (hasJobs) return alert("Bu müşteriye ait işler var. Önce o işleri silmelisin."); if (!confirm(`${c.company} müşterisini silmek istiyor musun?`)) return; await supabase.from("customers").delete().eq("id", c.id); await loadAll() }} className="bg-red-600 text-white px-3 py-2 rounded">Sil</button></td></tr>)}</tbody></table></Panel></div></div> }
function Column({ status, jobs, isAdmin, nextJob, prevJob, archiveJob, copyJob, setEditJob, makeJobPdf, makeDeliveryPdf }: any) { return <div className="bg-white border border-t-4 border-t-blue-600 rounded-xl p-3 min-h-[560px]"><div className="flex justify-between mb-3"><h3 className="font-black text-sm">{statusTitle[status as Status]}</h3><span>{jobs.length}</span></div><div className="space-y-3">{jobs.map((job: Job) => <div key={job.id} className={`border rounded-lg p-3 shadow-sm ${job.priority === "urgent" ? "bg-red-50 border-red-300" : "bg-white"}`}><button onClick={() => makeJobPdf(job)} className="text-xs text-blue-600 font-bold hover:underline">{jobNo(job)}</button><div className="font-black mt-2">{job.customer_name}</div><div className="text-sm">{job.job_name}</div><div className="text-sm mt-2">Adet: {job.quantity}</div><div className="text-sm">Teslim: {job.delivered}</div><div className="text-sm">Kalan: {Math.max(Number(job.quantity || 0) - Number(job.delivered || 0), 0)}</div>{job.priority === "urgent" && <div className="text-red-600 text-xs font-black">ACİL</div>}{isAdmin && <div className="text-sm font-bold">Kâr: {Number(job.profit || 0).toLocaleString("tr-TR")} ₺</div>}<div className="grid grid-cols-2 gap-2 mt-3">{status !== "printing" && <button onClick={() => prevJob(job)} className="bg-slate-200 py-2 rounded text-xs font-bold">← Geri</button>}{status !== "finished" && <button onClick={() => nextJob(job)} className="bg-blue-600 text-white py-2 rounded text-xs font-bold">İleri →</button>}</div>{Number(job.delivered || 0) > 0 && <button onClick={() => makeDeliveryPdf(job)} className="w-full mt-2 bg-green-600 text-white py-2 rounded text-xs font-bold">Teslim Fişi</button>}{isAdmin && <><button onClick={() => setEditJob(job)} className="w-full mt-2 bg-yellow-100 text-yellow-800 py-2 rounded text-xs font-bold">Düzenle</button><button onClick={() => copyJob(job)} className="w-full mt-2 bg-slate-900 text-white py-2 rounded text-xs font-bold">Kopyala</button><button onClick={() => archiveJob(job)} className="w-full mt-2 bg-red-100 text-red-700 py-2 rounded text-xs font-bold">Arşive Taşı</button></>}<details className="mt-2 text-xs"><summary className="cursor-pointer font-bold">Geçmiş</summary>{(job.logs || []).map((l, i) => <div key={i} className="border-t py-1">{new Date(l.created_at).toLocaleString("tr-TR")} - {l.text}</div>)}</details></div>)}</div></div> }
function JobsTable({ jobs, isAdmin, makeJobPdf, showCosts }: any) { return <table className="w-full text-sm"><thead><tr className="border-b text-left text-slate-500"><th className="py-2">İş No</th><th>Müşteri</th><th>İş</th><th>Adet</th><th>Teslim</th>{isAdmin && <th>Fiyat</th>}{showCosts && <><th>Maliyet</th><th>Kâr</th></>}<th>Durum</th>{isAdmin && <th>Fatura</th>}</tr></thead><tbody>{jobs.map((j: Job) => <tr key={j.id} className="border-b"><td className="py-3 font-bold"><button onClick={() => makeJobPdf(j)} className="text-blue-600 hover:underline">{jobNo(j)}</button></td><td>{j.customer_name}</td><td>{j.job_name}</td><td>{j.quantity}</td><td>{j.delivered}</td>{isAdmin && <td>{Number(j.price || 0).toLocaleString("tr-TR")} ₺</td>}{showCosts && <><td>{Number(j.total_cost || 0).toLocaleString("tr-TR")} ₺</td><td>{Number(j.profit || 0).toLocaleString("tr-TR")} ₺</td></>}<td>{statusTitle[j.status]}</td>{isAdmin && <td>{invoiceTitle[j.invoice_status]}</td>}</tr>)}</tbody></table> }
function InvoiceTable({ jobs, updateInvoice, makeJobPdf }: any) { return <table className="w-full text-sm"><thead><tr className="border-b text-left text-slate-500"><th className="py-2">İş No</th><th>Müşteri</th><th>İş</th><th>Tutar</th><th>Durum</th><th>İşlem</th></tr></thead><tbody>{jobs.map((j: Job) => <tr key={j.id} className="border-b"><td className="py-3 font-bold"><button onClick={() => makeJobPdf(j)} className="text-blue-600 hover:underline">{jobNo(j)}</button></td><td>{j.customer_name}</td><td>{j.job_name}</td><td>{Number(j.price || 0).toLocaleString("tr-TR")} ₺</td><td>{invoiceTitle[j.invoice_status]}</td><td className="space-x-2"><button onClick={() => updateInvoice(j, "waiting")} className="bg-slate-200 px-3 py-2 rounded">Bekliyor</button><button onClick={() => updateInvoice(j, "invoiced")} className="bg-green-600 text-white px-3 py-2 rounded">Kesildi</button><button onClick={() => updateInvoice(j, "paid")} className="bg-blue-600 text-white px-3 py-2 rounded">Ödendi</button></td></tr>)}</tbody></table> }
function Stocks({ stocks, stockForm, setStockForm, addStock, moveStock }: any) { return <div className="grid grid-cols-1 md:grid-cols-3 gap-5"><Panel><h2 className="font-black text-xl mb-4">Stok Ekle</h2><div className="space-y-3"><select value={stockForm.name} onChange={(e) => setStockForm({ ...stockForm, name: e.target.value })} className="border rounded-lg p-2 text-sm w-full"><option value="">Stok Seç</option><option value="1. Hamur">1. Hamur</option><option value="2. Hamur">2. Hamur</option><option value="Kuşe">Kuşe</option><option value="Bristol">Bristol</option><option value="Solvent">Solvent</option><option value="Alkol">Alkol</option><option value="Boya">Boya</option><option value="Sprey Gum">Sprey Gum</option><option value="Tutkal">Tutkal</option><option value="Selefon">Selefon</option><option value="Laminasyon Film">Laminasyon Film</option><option value="Kalıp">Kalıp</option><option value="Mürekkep">Mürekkep</option><option value="Koli">Koli</option><option value="Shrink">Shrink</option></select><select value={stockForm.type} onChange={(e) => setStockForm({ ...stockForm, type: e.target.value })} className="border rounded-lg p-2 text-sm w-full"><option value="">Tür Seç</option><option value="Kağıt">Kağıt</option><option value="Kimyasal">Kimyasal</option><option value="Boya">Boya</option><option value="Ambalaj">Ambalaj</option><option value="Baskı Malzemesi">Baskı Malzemesi</option><option value="Diğer">Diğer</option></select><Input p="Miktar" type="number" v={stockForm.quantity} c={(v: string) => setStockForm({ ...stockForm, quantity: v })} /><Input p="Birim" v={stockForm.unit} c={(v: string) => setStockForm({ ...stockForm, unit: v })} /><Input p="Minimum" type="number" v={stockForm.min_quantity} c={(v: string) => setStockForm({ ...stockForm, min_quantity: v })} /><button onClick={addStock} className="bg-blue-600 text-white px-5 py-3 rounded-lg font-bold">Stok Ekle</button></div></Panel><div className="md:col-span-2"><Panel><table className="w-full text-sm min-w-[650px]"><thead><tr className="border-b text-left text-slate-500"><th>Ad</th><th>Tür</th><th>Miktar</th><th>Min</th><th>İşlem</th></tr></thead><tbody>{stocks.map((s: Stock) => <tr key={s.id} className="border-b"><td className="py-3 font-bold">{s.name}</td><td>{s.type}</td><td className={Number(s.quantity) <= Number(s.min_quantity) ? "text-red-600 font-bold" : ""}>{s.quantity} {s.unit}</td><td>{s.min_quantity}</td><td className="space-x-2"><button onClick={() => moveStock(s, "Giriş")} className="bg-green-600 text-white px-3 py-2 rounded">Giriş</button><button onClick={() => moveStock(s, "Çıkış")} className="bg-red-600 text-white px-3 py-2 rounded">Çıkış</button></td></tr>)}</tbody></table></Panel></div></div> }
function CostSettingsPanel({ settings, setSettings, save }: any) { if (!settings) return <Panel>Maliyet ayarları yükleniyor...</Panel>; const set = (key: string, value: string) => setSettings({ ...settings, [key]: Number(value || 0) }); return <Panel><h2 className="text-xl font-black mb-4">Maliyet Ayarları</h2><div className="grid grid-cols-4 gap-3"><Input p="Kağıt Birim Fiyat" type="number" v={settings.paper_kg_price} c={(v: string) => set("paper_kg_price", v)} /><Input p="Baskı Forma Fiyatı" type="number" v={settings.print_form_price} c={(v: string) => set("print_form_price", v)} /><Input p="Kırım / Katlama Forma" type="number" v={settings.folding_form_price} c={(v: string) => set("folding_form_price", v)} /><Input p="Kesim / Diğer" type="number" v={settings.cutting_price} c={(v: string) => set("cutting_price", v)} /><Input p="Amerikan Cilt / Adet" type="number" v={settings.american_binding_price} c={(v: string) => set("american_binding_price", v)} /><Input p="İplik Dikiş / Adet" type="number" v={settings.thread_binding_price} c={(v: string) => set("thread_binding_price", v)} /><Input p="Tel Dikiş / Adet" type="number" v={settings.staple_binding_price} c={(v: string) => set("staple_binding_price", v)} /><Input p="Spiral / Adet" type="number" v={settings.spiral_binding_price} c={(v: string) => set("spiral_binding_price", v)} /><Input p="Kapak Baskı" type="number" v={settings.cover_print_price} c={(v: string) => set("cover_print_price", v)} /><Input p="Laminasyon / Adet" type="number" v={settings.lamination_price} c={(v: string) => set("lamination_price", v)} /><Input p="Fire %" type="number" v={settings.waste_percent} c={(v: string) => set("waste_percent", v)} /><Input p="Kâr %" type="number" v={settings.profit_percent} c={(v: string) => set("profit_percent", v)} /><Input p="KDV %" type="number" v={settings.vat_percent} c={(v: string) => set("vat_percent", v)} /></div><button onClick={save} className="mt-5 bg-blue-600 text-white px-6 py-3 rounded-lg font-bold">Ayarları Kaydet</button><p className="text-sm text-slate-500 mt-4">Yeni İş Girişi ekranında “Maliyeti Hesapla” butonu bu ayarları kullanır.</p></Panel> }
function Deliveries({ jobs }: any) { const rows = jobs.flatMap((j: Job) => (j.deliveries || []).map((d) => [jobNo(j), j.customer_name, j.job_name, String(d.amount), d.note, new Date(d.created_at).toLocaleString("tr-TR")])); return <SimpleTable head={["İş No", "Müşteri", "İş", "Adet", "Not", "Tarih"]} rows={rows} /> }
function DeliveryModal({ job, form, setForm, close, save }: any) { const remaining = Math.max(Number(job.quantity || 0) - Number(job.delivered || 0), 0); const after = Math.max(remaining - Number(form.amount || 0), 0); return <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50"><div className="bg-white rounded-2xl w-[520px] shadow-2xl overflow-hidden"><div className="bg-[#071d35] text-white p-5"><h2 className="text-2xl font-black">Teslimat Girişi</h2><div className="text-sm opacity-80">{job.customer_name} - {job.job_name}</div></div><div className="p-6 space-y-4"><div className="grid grid-cols-3 gap-3"><div className="bg-slate-100 rounded-xl p-4"><div className="text-xs text-slate-500 font-bold">Toplam</div><div className="text-2xl font-black">{job.quantity}</div></div><div className="bg-slate-100 rounded-xl p-4"><div className="text-xs text-slate-500 font-bold">Teslim</div><div className="text-2xl font-black">{job.delivered}</div></div><div className="bg-slate-100 rounded-xl p-4"><div className="text-xs text-slate-500 font-bold">Kalan</div><div className="text-2xl font-black">{remaining}</div></div></div><Input p="Teslim edilen adet" type="number" v={form.amount} c={(v: string) => setForm({ ...form, amount: v })} /><textarea className="border rounded-lg p-3 w-full min-h-[110px]" placeholder="Teslimat notu" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /><div className="bg-blue-50 text-blue-800 rounded-xl p-4 font-bold">Bu teslimattan sonra kalan: {after}</div><div className="flex gap-3"><button onClick={close} className="flex-1 bg-slate-200 py-3 rounded-lg font-bold">İptal</button><button onClick={save} className="flex-1 bg-blue-600 text-white py-3 rounded-lg font-bold">Teslim Et</button></div></div></div></div> }
function EditModal({ job, setJob, save }: any) { return <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50"><div className="bg-white rounded-xl p-6 w-[700px] max-h-[90vh] overflow-auto"><h2 className="text-xl font-black mb-4">İş Düzenle</h2><div className="grid grid-cols-2 gap-3"><Input p="İş Adı" v={job.job_name} c={(v: string) => setJob({ ...job, job_name: v })} /><Input p="Adet" type="number" v={job.quantity} c={(v: string) => setJob({ ...job, quantity: Number(v) })} /><Input p="Fiyat" type="number" v={job.price} c={(v: string) => setJob({ ...job, price: Number(v) })} /><Input p="Kağıt Maliyeti" type="number" v={job.paper_cost} c={(v: string) => setJob({ ...job, paper_cost: Number(v) })} /><Input p="Baskı Maliyeti" type="number" v={job.print_cost} c={(v: string) => setJob({ ...job, print_cost: Number(v) })} /><Input p="Cilt Maliyeti" type="number" v={job.binding_cost} c={(v: string) => setJob({ ...job, binding_cost: Number(v) })} /><Input p="Laminasyon Maliyeti" type="number" v={job.lamination_cost} c={(v: string) => setJob({ ...job, lamination_cost: Number(v) })} /><Input p="İşçilik" type="number" v={job.labor_cost} c={(v: string) => setJob({ ...job, labor_cost: Number(v) })} /><Input p="Teslim Tarihi" type="date" v={job.deadline} c={(v: string) => setJob({ ...job, deadline: v })} /><Input p="Not" v={job.note} c={(v: string) => setJob({ ...job, note: v })} /></div><div className="flex gap-3 mt-5"><button onClick={save} className="bg-blue-600 text-white px-5 py-3 rounded-lg font-bold">Kaydet</button><button onClick={() => setJob(null)} className="bg-slate-200 px-5 py-3 rounded-lg font-bold">Kapat</button></div></div></div> }
function SimpleTable({ head, rows }: any) { return <table className="w-full text-sm"><thead><tr className="border-b text-left text-slate-500">{head.map((h: string) => <th key={h} className="py-2">{h}</th>)}</tr></thead><tbody>{rows.map((r: any[], i: number) => <tr key={i} className="border-b">{r.map((x, k) => <td key={k} className="py-3">{x}</td>)}</tr>)}</tbody></table> }
