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

    const admin = createClient(supabaseUrl, serviceKey)

    const authHeader = req.headers.get("Authorization")
    if (!authHeader) {
      return new Response(JSON.stringify({ error: "Yetkisiz giriş" }), { status: 401, headers })
    }

    const token = authHeader.replace("Bearer ", "")
    const { data: userData, error: userError } = await admin.auth.getUser(token)

    if (userError || !userData.user) {
      return new Response(JSON.stringify({ error: "Geçersiz oturum" }), { status: 401, headers })
    }

    const { data: adminProfile } = await admin
      .from("profiles")
      .select("role")
      .eq("id", userData.user.id)
      .single()

    if (adminProfile?.role !== "admin") {
      return new Response(JSON.stringify({ error: "Sadece admin işlem yapabilir" }), { status: 403, headers })
    }

    const body = await req.json()
    const { username, password, full_name, role, department } = body

    if (!username || !password) {
      return new Response(JSON.stringify({ error: "Kullanıcı adı ve şifre zorunlu" }), { status: 400, headers })
    }

    const cleanUsername = String(username).trim().toLowerCase()
    const email = cleanUsername.includes("@") ? cleanUsername : `${cleanUsername}@balbasim.com`

    const { data, error } = await admin.auth.admin.createUser({
      email,
      password,
      email_confirm: true,
      user_metadata: { full_name: full_name || cleanUsername },
    })

    if (error) {
      return new Response(JSON.stringify({ error: error.message }), { status: 400, headers })
    }

    const { error: profileError } = await admin.from("profiles").upsert({
      id: data.user.id,
      email,
      username: email.split("@")[0],
      full_name: full_name || cleanUsername,
      role: role || "staff",
      department: department || "Personel",
      active: true,
    })

    if (profileError) {
      return new Response(JSON.stringify({ error: profileError.message }), { status: 400, headers })
    }

    return new Response(JSON.stringify({ ok: true, user: data.user }), { headers })
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e)
    return new Response(JSON.stringify({ error: message }), { status: 500, headers })
  }
})
