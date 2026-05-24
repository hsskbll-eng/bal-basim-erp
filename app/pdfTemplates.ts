function jobNo(job: any) {
  const year = String(new Date(job.created_at).getFullYear()).slice(2)

  return `${String(job.id).padStart(4, "0")}-${year}`
}
function clean(v: any) {
  return String(v ?? "-")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
}

export function openJobPrint(job: any) {
  const w = window.open("", "_blank")
  if (!w) return alert("Açılır pencere engellendi.")

  w.document.write(`
<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8" />
<title>${jobNo(job)}</title>
<style>
.logoBox{
  text-align:center;
  margin-bottom:8px;
}

.logoBox img{
  width:260px;
  height:auto;
}  
body{margin:0;background:#d1d5db;font-family:Arial,Helvetica,sans-serif;color:#0b1f35}
  .page{width:210mm;min-height:297mm;background:#fff;margin:18px auto;padding:13mm}
  .date{font-size:12px;margin-bottom:8px}
  .brand{text-align:center;font-size:34px;font-weight:900;letter-spacing:-1px}
  .sub{text-align:center;font-size:17px;font-weight:700;margin:6px 0 18px}
  table{width:100%;border-collapse:collapse}
  td,th{border:1px solid #9ca3af;padding:9px 10px;font-size:13px}
  .dark{background:#071d35;color:white;font-weight:900}
  .jobno{font-size:18px;font-weight:900;text-align:center}
  .section{background:#071d35;color:white;text-align:center;font-size:18px;font-weight:900}
  .icon{width:36px;text-align:center;font-size:18px}
  .name{font-weight:900;width:130px}
  .notes{height:285px;vertical-align:top}
  .noteLine{height:34px;border-bottom:1px dotted #6b7280}
  .signs{display:grid;grid-template-columns:1fr 1fr 1fr;border:1px solid #9ca3af;margin-top:22px}
  .sig{text-align:center;height:115px;border-right:1px solid #9ca3af;padding-top:18px;font-weight:900}
  .sig:last-child{border-right:0}
  .sigIcon{font-size:28px;margin-bottom:8px}
  .line{width:70%;border-bottom:2px dotted #111;margin:28px auto 0}
  .footer{text-align:center;font-size:12px;margin-top:14px}
  @media print{body{background:white}.page{margin:0}}
</style>
</head>
<body>
<div class="page">
  <div class="date">${new Date().toLocaleString("tr-TR")}</div>
 <div class="logoBox">
  <img src="/logo.png" />
</div>
  <div class="sub">Matbaa İş Emri Formu</div>

  <table>
    <tr>
      <td class="dark">Firma Adı</td>
      <td>${clean(job.customer_name)}</td>
      <td class="jobno" rowspan="4">İş Emri No<br><br>${jobNo(job)}</td>
    </tr>
    <tr><td class="dark">Tarih</td><td>${new Date(job.created_at).toLocaleString("tr-TR")}</td></tr>
    <tr><td class="dark">İşin Adı</td><td>${clean(job.job_name)}</td></tr>
    <tr><td class="dark">Emri Oluşturan</td><td>BAL BASIM YAYIN</td></tr>
  </table>

  <br/>

  <table>
    <tr>
      <th colspan="3" class="section">İÇ</th>
      <th colspan="3" class="section">KAPAK</th>
    </tr>

    <tr>
      <td class="icon">▦</td><td class="name">Adet</td><td>${clean(job.quantity)}</td>
      <td class="icon">□</td><td class="name">Kağıt Cinsi</td><td>${clean(job.cover_type || job.cover_paper_type)}</td>
    </tr>

    <tr>
      <td class="icon">□</td><td class="name">Sayfa Ölçüsü</td><td>${clean(job.size)}</td>
      <td class="icon">▧</td><td class="name">Selefon</td><td>${clean(job.cellophane)}</td>
    </tr>

    <tr>
      <td class="icon">▤</td><td class="name">Sayfa Sayısı</td><td>${clean(job.page_count)}</td>
      <td class="icon">▣</td><td class="name">Kapak İçi Baskı</td><td>${clean(job.cover_inside_print)}</td>
    </tr>

    <tr>
      <td class="icon">●</td><td class="name">Renk Bilgisi</td><td>${clean(job.color)}</td>
      <td class="icon">▣</td><td class="name">Kapak Gramaj</td><td>${clean(job.cover_gram)}</td>
    </tr>

    <tr>
      <td class="icon">▤</td><td class="name">Kağıt Cinsi</td><td>${clean(job.inner_paper_type)}</td>
      <td class="icon">□</td><td class="name">Kapak Ebat</td><td>${clean(job.cover_size)}</td>
    </tr>

    <tr>
      <td class="icon">G</td><td class="name">Kağıt Gramajı</td><td>${clean(job.inner_paper_gram)}</td>
      <td class="icon">□</td><td class="name">Kapak Miktar</td><td>${clean(job.cover_amount)}</td>
    </tr>

    <tr>
      <td class="icon">▣</td><td class="name">Kapak Kağıdı</td><td>${clean(job.cover_type)}</td>
      <td class="icon">▧</td><td class="name">Laminasyon</td><td>${clean(job.lamination)}</td>
    </tr>

    <tr>
      <td class="icon">≋</td><td class="name">Cilt Şekli</td><td>${clean(job.binding)}</td>
      <td rowspan="5" colspan="3" class="notes">
        <b>Notlar</b>
        <div class="noteLine"></div>
        <div class="noteLine"></div>
        <div class="noteLine"></div>
        <div class="noteLine"></div>
        <div class="noteLine"></div>
        <br/>
        ${clean(job.note)}
      </td>
    </tr>

    <tr><td class="icon">▭</td><td class="name">Baskı Tipi</td><td>${clean(job.print_type)}</td></tr>
    <tr><td class="icon">▦</td><td class="name">Shrink Adedi</td><td>${clean(job.shrink_amount)}</td></tr>
    <tr><td class="icon">🚚</td><td class="name">Teslim Tarihi</td><td>${clean(job.deadline)}</td></tr>
  </table>

  <div class="signs">
    <div class="sig"><div class="sigIcon">👤</div>HAZIRLAYAN<div class="line"></div></div>
    <div class="sig"><div class="sigIcon">✓</div>BASKI ONAY<div class="line"></div></div>
    <div class="sig"><div class="sigIcon">🚚</div>TESLİM ONAY<div class="line"></div></div>
  </div>

  <div class="footer">Not: Üretimden önce tüm bilgiler kontrol edilmelidir.</div>
</div>

<script>
  setTimeout(() => window.print(), 500)
</script>
</body>
</html>
  `)

  w.document.close()
}