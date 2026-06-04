import { serve } from "https://deno.land/std@0.224.0/http/server.ts"
import { createClient } from "https://esm.sh/@supabase/supabase-js@2"

serve(async (req) => {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
    "Content-Type": "application/json",
  }

  if (req.method === "OPTIONS") return new Response("ok", { headers })

  try {
    const supabaseUrl = Deno.env.get("SUPABASE_URL")
    const serviceKey = Deno.env.get("SERVICE_ROLE_KEY")

    if (!supabaseUrl || !serviceKey) {
      return new Response(JSON.stringify({ error: "Server ayarları eksik" }), { status: 500, headers })
    }

    const authHeader = req.headers.get("Authorization")
    if (!authHeader) {
      return new Response(JSON.stringify({ error: "Yetkisiz giriş" }), { status: 401, headers })
    }

    const admin = createClient(supabaseUrl, serviceKey)
    const token = authHeader.replace("Bearer ", "")
    const { data: userData, error: userError } = await admin.auth.getUser(token)

    if (userError || !userData.user) {
      return new Response(JSON.stringify({ error: "Geçersiz oturum" }), { status: 401, headers })
    }

    const { data: profile, error: profileError } = await admin
      .from("profiles")
      .select("role, active")
      .eq("id", userData.user.id)
      .single()

    if (profileError || profile?.role !== "admin" || profile?.active === false) {
      return new Response(JSON.stringify({ error: "Sadece aktif admin personel oluşturabilir" }), { status: 403, headers })
    }

    const body = await req.json()
    const { username, password, full_name, role, department } = body

    if (!username || !password) {
      return new Response(JSON.stringify({ error: "Kullanıcı adı ve şifre zorunlu" }), { status: 400, headers })
    }

    if (String(password).length < 4) {
      return new Response(JSON.stringify({ error: "Şifre en az 4 karakter olmalı" }), { status: 400, headers })
    }

    const cleanUsername = String(username).trim().toLowerCase().replace(/\s+/g, "")
    const email = `${cleanUsername}@balbasim.com`

    const { data, error } = await admin.auth.admin.createUser({
      email,
      password,
      email_confirm: true,
      user_metadata: { full_name: full_name || cleanUsername },
    })

    if (error) {
      return new Response(JSON.stringify({ error: error.message }), { status: 400, headers })
    }

    const { error: upsertError } = await admin.from("profiles").upsert({
      id: data.user.id,
      email,
      username: cleanUsername,
      full_name: full_name || cleanUsername,
      role: role || "staff",
      department: department || "Personel",
      active: true,
    })

    if (upsertError) {
      return new Response(JSON.stringify({ error: upsertError.message }), { status: 400, headers })
    }

    return new Response(JSON.stringify({ ok: true, user: { id: data.user.id, email } }), { headers })
  } catch (e) {
    return new Response(JSON.stringify({ error: String(e?.message || e) }), { status: 500, headers })
  }
})
