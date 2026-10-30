// Datos de conexión al proyecto de Supabase (Project Settings › API).
// La clave "anon" es pública por diseño: los datos los protege la seguridad por filas (RLS).
window.CONFIG = {
  SUPABASE_URL: "https://hfqbbxcswllntphswtdn.supabase.co",
  SUPABASE_ANON_KEY: "sb_publishable_c9fpBV7aC1BkbvHJJgVb9w_H6Vyk7rH",
  COMISION_PORCENTAJE: 3,
  DIAS_CREDITO: 7,
  DIA_TASA_COMISION: 29,
  // Datos de la empresa para la nota de entrega y los reportes en PDF.
  // LOGO: ruta de la imagen del logo (ej. "icons/logo.png"); vacío = logo en texto.
  EMPRESA: {
    NOMBRE: "INVERSIONES CHOCOMUNDO 2030, C.A.",
    MARCA: "Choco tentación",
    RIF: "J-505733605",
    DIRECCION: "Calle de Acceso, local Centro Empresarial Caribe Nro. Galpón G-10, Sector Asentamiento Campesino Morita 1, Parcela 10 Lota A-1, Turmero Edo. Aragua, Zona Postal 2115",
    TELEFONO: "0416-642.76.80",
    LOGO: ""
  }
};
