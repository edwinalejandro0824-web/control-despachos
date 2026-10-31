// Datos de conexión al proyecto de Supabase (Project Settings › API).
// La clave "anon" es pública por diseño: los datos los protege la seguridad por filas (RLS).
window.CONFIG = {
  SUPABASE_URL: "https://hfqbbxcswllntphswtdn.supabase.co",
  SUPABASE_ANON_KEY: "sb_publishable_c9fpBV7aC1BkbvHJJgVb9w_H6Vyk7rH",
  COMISION_PORCENTAJE: 3,
  DIAS_CREDITO: 7,
  DIA_TASA_COMISION: 29,
  IVA_PORCENTAJE: 16,
  // Catálogo de distribución (Brochure de venta 2026). Precios en USD SIN IVA.
  // Para cambiar un precio o una descripción, edite la línea del producto.
  // Con venta en Bs., el precio se convierte con la tasa BCV de la fecha del despacho.
  CATALOGO: {
    "CJ":      { desc: "Brownie Clásico (caja de 6 UND)", precio: 3.50 },
    "KG":      { desc: "Brownie Clásico por KG", precio: 8.00 },
    "MINI":    { desc: "Brownie Bite (UND de 3 cm x 3 cm)", precio: 0.20 },
    "GALLETAS":{ desc: "Galletas de Brownie (paquete de 6 UND)", precio: 1.00 },
    "PMV":     { desc: "Pre-Mezcla de Vainilla 400 gr", precio: 1.50 },
    "PMCH":    { desc: "Pre-Mezcla de Chocolate 400 gr", precio: 1.98 },
    "PMB":     { desc: "Pre-Mezcla de Brownie 400 gr", precio: 2.23 },
    "PMS":     { desc: "Pre-Mezcla Surtida 400 gr", precio: 1.9033 },
    "KG PMB":  { desc: "Pre-Mezcla de Brownie (saco, por KG)", precio: 5.40 },
    "KG PMCH": { desc: "Pre-Mezcla de Chocolate (saco, por KG)", precio: 4.80 },
    "KG PMV":  { desc: "Pre-Mezcla de Vainilla (saco, por KG)", precio: 3.50 },
    "MHCH":    { desc: "Pre-Mezcla de Helado de Máquina 550 gr · Chocolate", precio: 7.00 },
    "MHM":     { desc: "Pre-Mezcla de Helado de Máquina 550 gr · Mantecado", precio: 7.00 }
  },
  // Datos de la empresa para la nota de entrega y los reportes en PDF.
  // LOGO: ruta de la imagen del logo (ej. "icons/logo.png"); vacío = logo en texto.
  EMPRESA: {
    NOMBRE: "INVERSIONES CHOCOMUNDO 2030, C.A.",
    MARCA: "Choco tentación",
    RIF: "J-505733605",
    DIRECCION: "Calle de Acceso, local Centro Empresarial Caribe Nro. Galpón G-10, Sector Asentamiento Campesino Morita 1, Parcela 10 Lota A-1, Turmero Edo. Aragua, Zona Postal 2115",
    TELEFONO: "0416-642.76.80",
    LOGO: "icons/logo.png"
  }
};
