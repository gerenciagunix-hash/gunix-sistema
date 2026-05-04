const { getRange, appendRow, updateCell, JOTFORM_API_KEY, JOTFORM_FORM_ID } = require('./sheets');
const https = require('https');

function calcM2Equiv(m2, m3) {
  m2 = parseFloat(m2) || 0; m3 = parseFloat(m3) || 0;
  return m2 + (Math.max(0, m3 - 12) * 2);
}
function calcJornales(m2Equiv) {
  if (m2Equiv >= 120) return 2;
  if (m2Equiv >= 90) return 1.5;
  return 1;
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  const { action } = req.query;
  try {
    if (req.method === 'GET') {
      if (action === 'dashboard') {
        const empresas = await getRange('Empresas');
        const remitos = await getRange('Remitos');
        const certs = await getRange('Certificados');
        const facturas = await getRange('Facturas');
        const hoy = new Date();
        const mes = hoy.getMonth();
        const anio = hoy.getFullYear();
        const remMes = remitos.filter(r => { const f = new Date(r.Fecha); return f.getMonth() === mes && f.getFullYear() === anio; });
        const pendiente = facturas.filter(f => f.Estado === 'Enviada' || f.Estado === 'Emitida').reduce((s, f) => s + (parseFloat(f.Total) || 0), 0);
        const vencidas = facturas.filter(f => f.Estado !== 'Cobrada' && new Date(f.FechaVencimiento) < hoy).reduce((s, f) => s + (parseFloat(f.Total) || 0), 0);
        return res.json({ totalClientes: empresas.filter(c => c.Estado === 'Activo').length, remitosDelMes: remMes.length, sinCertificar: remitos.filter(r => !r.EstadoCert || r.EstadoCert === 'Pendiente').length, certificados: certs.length, pendienteCobro: pendiente, vencidas, ultimosRemitos: remMes.slice(-5).reverse(), facturasPendientes: facturas.filter(f => f.Estado !== 'Cobrada').slice(-5) });
      }
      if (action === 'clientes') {
        const empresas = await getRange('Empresas');
        const obras = await getRange('Obras');
        const remitos = await getRange('Remitos');
        return res.json(empresas.map(e => {
          const obrasEmp = obras.filter(o => o.EmpresaID === e.ID);
          const remEmp = remitos.filter(r => r.ClienteID === e.ID);
          const obra1 = obrasEmp[0] || {};
          return { ...e, Empresa: e.RazonSocial, Obra: obra1.NombreObra || '', Direccion: obra1.DireccionObra || e.DirFiscal || '', obras: obrasEmp, totalRemitos: remEmp.length, sinCertificar: remEmp.filter(r => !r.EstadoCert || r.EstadoCert === 'Pendiente').length, remitos: remEmp };
        }));
      }
      if (action === 'empleados') return res.json(await getRange('Empleados'));
      if (action === 'remitos') return res.json(await getRange('Remitos'));
      if (action === 'certificados') return res.json(await getRange('Certificados'));
      if (action === 'facturas') return res.json(await getRange('Facturas'));
      if (action === 'jornales') {
        const { desde, hasta } = req.query;
        const jornales = await getRange('Jornales');
        const empleados = await getRange('Empleados');
        const periodo = jornales.filter(j => { const f = String(j.Fecha || '').slice(0, 10); return f >= desde && f <= hasta; });
        const liquidacion = empleados.filter(e => e.Activo === 'Si').map(e => {
          const remitosEmp = periodo.filter(j => (j.Personal || '').split(', ').some(p => p.trim() === e.NombreJotForm));
          const dias = new Set(remitosEmp.map(r => String(r.Fecha || '').slice(0, 10)));
          const totalJornales = remitosEmp.reduce((s, r) => s + (parseFloat(r.Jornales) || 1), 0);
          const totalComida = dias.size * (parseFloat(e.ComidaDia) || 0);
          const totalJornal = totalJornales * (parseFloat(e.JornadaBase) || 0);
          return { id: e.ID, nombre: e.NombreJotForm, alias: e.Alias, cuadrilla: e.Cuadrilla, rol: e.Rol, jornadaBase: parseFloat(e.JornadaBase) || 0, comidaDia: parseFloat(e.ComidaDia) || 0, jornales: Math.round(totalJornales * 10) / 10, diasTrabajados: dias.size, servicios: remitosEmp.length, totalJornal: Math.round(totalJornal), totalComida: Math.round(totalComida), total: Math.round(totalJornal + totalComida), detalle: remitosEmp.map(r => ({ fecha: String(r.Fecha || '').slice(0, 10), empresa: r.Empresa, m2: r.M2, m3: r.M3, m2Equiv: r.M2Equiv, jornales: r.Jornales })) };
        });
        return res.json({ ok: true, desde, hasta, empleados: liquidacion, granTotal: liquidacion.reduce((s, e) => s + e.total, 0) });
      }
    }
    if (req.method === 'POST') {
      const body = req.body;
      if (action === 'editarEmpresa') {
        const empresas = await getRange('Empresas');
        const idx = empresas.findIndex(e => e.ID === body.ID);
        if (idx < 0) return res.json({ ok: false, error: 'No encontrada' });
        const headers = ['ID', 'RazonSocial', 'CUIT', 'CondicionIVA', 'DirFiscal', 'Localidad', 'TelEmpresa', 'EmailEmpresa', 'ContactoGeneral', 'CobrEmail', 'CobrTel', 'CobrNombre', 'Estado', 'TotalServicios'];
        const campos = { RazonSocial: body.RazonSocial, CUIT: body.CUIT, CondicionIVA: body.CondicionIVA, DirFiscal: body.DirFiscal, Localidad: body.Localidad, TelEmpresa: body.TelEmpresa, EmailEmpresa: body.EmailEmpresa, CobrNombre: body.CobrNombre, CobrEmail: body.CobrEmail, CobrTel: body.CobrTel };
        for (const [key, val] of Object.entries(campos)) { const col = headers.indexOf(key); if (col >= 0 && val !== undefined) await updateCell('Empresas', idx + 1, col, val); }
        return res.json({ ok: true });
      }
      if (action === 'agregarObra') {
        const id = 'OBRA_' + Date.now();
        await appendRow('Obras', [id, body.EmpresaID, body.RazonSocial, body.NombreObra, body.DireccionObra || '', body.LocalidadObra || '', body.ObraNombre || '', body.ObraEmail || '', body.ObraTel || '', body.Presupuesto || '', body.PrecioFyF || '', body.PrecioFrat || '', body.PrecioReg || '', body.DifEspesor || '', body.Minimo || '', body.CACBase || '', 'Activo', 0]);
        return res.json({ ok: true, id });
      }
    }
    res.status(400).json({ error: 'Accion no reconocida: ' + action });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
};
