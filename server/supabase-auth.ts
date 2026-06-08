import { ENV } from "./_core/env";

type SupabaseAuthUser = {
  id: string;
  email?: string | null;
  user_metadata?: {
    name?: string | null;
    full_name?: string | null;
  } | null;
};

type SupabaseAuthResponse = {
  user?: SupabaseAuthUser | null;
};

type SupabaseAuthInput = {
  email: string;
  password: string;
  name?: string;
};

function getSupabaseConfig() {
  const url = ENV.supabaseUrl.replace(/\/$/, "");
  const apiKey = ENV.supabaseServiceRoleKey || ENV.supabaseAnonKey;

  if (!url || !apiKey) {
    throw new Error("Supabase não configurado no servidor");
  }

  return { url, apiKey };
}

async function requestSupabaseAuth(path: string, body: Record<string, unknown>) {
  const { url, apiKey } = getSupabaseConfig();
  const response = await fetch(`${url}/auth/v1/${path}`, {
    method: "POST",
    headers: {
      apikey: apiKey,
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  const data = await response.json().catch(() => ({}));

  if (!response.ok) {
    const message =
      typeof data?.msg === "string"
        ? data.msg
        : typeof data?.message === "string"
          ? data.message
          : "Falha na autenticação com Supabase";
    throw new Error(message);
  }

  return data as SupabaseAuthResponse;
}

function normalizeSupabaseUser(data: SupabaseAuthResponse) {
  const user = data.user;

  if (!user?.id) {
    throw new Error("Supabase não retornou usuário");
  }

  const metadataName = user.user_metadata?.name ?? user.user_metadata?.full_name ?? null;
  const name = metadataName || user.email || "Usuário Saúde de Atleta";

  return {
    openId: `supabase:${user.id}`,
    email: user.email ?? null,
    name,
  };
}

export async function signUpWithSupabase({ email, password, name }: SupabaseAuthInput) {
  const data = await requestSupabaseAuth("signup", {
    email,
    password,
    data: {
      name: name || email,
    },
  });

  return normalizeSupabaseUser(data);
}

export async function signInWithSupabase({ email, password }: SupabaseAuthInput) {
  const data = await requestSupabaseAuth("token?grant_type=password", {
    email,
    password,
  });

  return normalizeSupabaseUser(data);
}
