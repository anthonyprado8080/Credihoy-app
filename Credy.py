from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
import pandas as pd
from datetime import datetime, timedelta
import uuid
import math
import io
import zipfile

app = Flask(__name__)
app.secret_key = "credihoy_seguridad_2026"

ARCHIVO_PRINCIPAL = "registro_prestamos.xlsx"
ARCHIVO_HISTORIAL = "historial_pagos.xlsx"

# ==========================================
# LÓGICA DE DATOS PROFESIONAL
# ==========================================
class GestorPrestamos:
    def obtener_clientes(self):
        if os.path.exists(ARCHIVO_PRINCIPAL):
            return pd.read_excel(ARCHIVO_PRINCIPAL).to_dict(orient='records')
        return []

    def obtener_cliente(self, codigo):
        if os.path.exists(ARCHIVO_PRINCIPAL):
            df = pd.read_excel(ARCHIVO_PRINCIPAL)
            cliente = df[df['Código'] == codigo]
            if not cliente.empty:
                return cliente.iloc[0].to_dict()
        return None

    def obtener_otros_prestamos(self, codigo_actual, dni):
        if os.path.exists(ARCHIVO_PRINCIPAL):
            df = pd.read_excel(ARCHIVO_PRINCIPAL)
            # Aseguramos que el DNI se lea como texto para evitar errores de búsqueda
            otros = df[(df['DNI'].astype(str) == str(dni)) & (df['Código'] != codigo_actual)]
            return otros.to_dict(orient='records')
        return []

    def registrar_prestamo(self, dni, nombres, monto, interes, modalidad, cuotas):
        dni = str(dni).strip()
        if len(dni) != 8 or not dni.isdigit():
            raise ValueError("El DNI debe contener exactamente 8 dígitos numéricos.")

        monto_str = str(monto).replace(',', '.')
        interes_str = str(interes).replace(',', '.')
        
        monto = float(monto_str)
        interes = float(interes_str)
        cuotas = int(cuotas)

        if cuotas <= 0:
            raise ValueError("El número de cuotas debe ser mayor a 0.")
        if monto <= 0:
            raise ValueError("El monto prestado debe ser mayor a 0.")

        monto_interes = monto * (interes / 100)
        total = round(monto + monto_interes, 2)
        
        # Redondeo peruano a monedas de 10 céntimos
        cuota_exacta = total / cuotas
        monto_cuota_base = math.ceil(cuota_exacta * 10) / 10
        
        codigo = f"CLI-{dni[-4:]}-{str(uuid.uuid4())[:4].upper()}"
        fecha = datetime.now().strftime("%Y-%m-%d")

        datos = {
            "Código": [codigo], "DNI": [dni], "Nombres": [nombres.title()],
            "Fecha": [fecha], "Monto (S/)": [float(monto)], "Interés (%)": [float(interes)],
            "Total (S/)": [float(total)], "Modalidad": [modalidad], 
            "Cuotas Total": [int(cuotas)], "Monto Cuota (S/)": [float(monto_cuota_base)],
            "Pagadas": [int(0)], "Saldo (S/)": [float(total)], "Estado": ["Activo"]
        }
        
        df_nuevo = pd.DataFrame(datos)
        if os.path.exists(ARCHIVO_PRINCIPAL):
            df_existente = pd.read_excel(ARCHIVO_PRINCIPAL)
            df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
            df_final.to_excel(ARCHIVO_PRINCIPAL, index=False)
        else:
            df_nuevo.to_excel(ARCHIVO_PRINCIPAL, index=False)

    def cobrar_cuota(self, codigo):
        if not os.path.exists(ARCHIVO_PRINCIPAL):
            return

        df = pd.read_excel(ARCHIVO_PRINCIPAL)
        
        # Forzar formatos para evitar el error "Invalid value for dtype int64"
        df['Saldo (S/)'] = df['Saldo (S/)'].astype(float)
        df['Monto Cuota (S/)'] = df['Monto Cuota (S/)'].astype(float)
        df['Pagadas'] = df['Pagadas'].astype(int)
        
        idx_lista = df[df['Código'] == codigo].index
        if len(idx_lista) == 0:
            return
            
        idx = idx_lista[0]
        
        if str(df.at[idx, 'Estado']).strip().upper() == "PAGADO":
            return
        
        saldo_actual = round(float(df.at[idx, 'Saldo (S/)']), 2)
        monto_cuota_base = round(float(df.at[idx, 'Monto Cuota (S/)']), 2)
        cuotas_totales = int(df.at[idx, 'Cuotas Total'])
        cuotas_pagadas = int(df.at[idx, 'Pagadas'])
        
        # Ajuste matemático de última cuota
        if (cuotas_pagadas + 1) >= cuotas_totales or saldo_actual <= monto_cuota_base + 0.10:
            pago_real = saldo_actual
            nuevo_saldo = 0.0
            df.at[idx, 'Estado'] = "Pagado"
        else:
            pago_real = monto_cuota_base
            nuevo_saldo = round(saldo_actual - pago_real, 2)
            
        df.at[idx, 'Pagadas'] = cuotas_pagadas + 1
        df.at[idx, 'Saldo (S/)'] = float(nuevo_saldo)
        
        # Seguro anti-céntimos fantasma
        if float(df.at[idx, 'Saldo (S/)']) <= 0.10:
            df.at[idx, 'Saldo (S/)'] = 0.0
            df.at[idx, 'Estado'] = "Pagado"
            
        df.to_excel(ARCHIVO_PRINCIPAL, index=False)

        fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df_hist = pd.DataFrame({
            "Código Cliente": [codigo], 
            "Fecha y Hora": [fecha_hora], 
            "Monto Pagado (S/)": [round(float(pago_real), 2)]
        })
        
        if os.path.exists(ARCHIVO_HISTORIAL):
            df_hist_existente = pd.read_excel(ARCHIVO_HISTORIAL)
            pd.concat([df_hist_existente, df_hist], ignore_index=True).to_excel(ARCHIVO_HISTORIAL, index=False)
        else:
            df_hist.to_excel(ARCHIVO_HISTORIAL, index=False)

    def obtener_historial(self, codigo):
        if os.path.exists(ARCHIVO_HISTORIAL):
            df = pd.read_excel(ARCHIVO_HISTORIAL)
            return df[df['Código Cliente'] == codigo].to_dict(orient='records')
        return []

    def eliminar_cliente(self, codigo):
        if os.path.exists(ARCHIVO_PRINCIPAL):
            df = pd.read_excel(ARCHIVO_PRINCIPAL)
            df = df[df['Código'] != codigo]
            df.to_excel(ARCHIVO_PRINCIPAL, index=False)
        if os.path.exists(ARCHIVO_HISTORIAL):
            df_h = pd.read_excel(ARCHIVO_HISTORIAL)
            df_h = df_h[df_h['Código Cliente'] != codigo]
            df_h.to_excel(ARCHIVO_HISTORIAL, index=False)

gestor = GestorPrestamos()

# ==========================================
# RUTAS WEB
# ==========================================
@app.route('/')
def inicio():
    # Buscador implementado
    query = request.args.get('q', '').strip().lower()
    clientes = gestor.obtener_clientes()
    
    if query:
        clientes_filtrados = []
        for c in clientes:
            if (query in str(c['DNI']).lower() or 
                query in str(c['Código']).lower() or 
                query in str(c['Nombres']).lower()):
                clientes_filtrados.append(c)
        clientes = clientes_filtrados
        
    return render_template('index.html', clientes=clientes, query=query)

@app.route('/registrar', methods=['POST'])
def registrar():
    try:
        gestor.registrar_prestamo(
            request.form['dni'], request.form['nombres'], request.form['monto'],
            request.form['interes'], request.form['modalidad'], request.form['cuotas']
        )
        flash("✅ Préstamo generado con éxito en el sistema.", "success")
    except ValueError as ve:
        flash(f"⚠️ {ve}", "danger")
    except Exception as e:
        flash(f"⚠️ Error inesperado: {e}", "danger")
    return redirect(url_for('inicio'))

@app.route('/cliente/<codigo>')
def detalle_cliente(codigo):
    cliente = gestor.obtener_cliente(codigo)
    if not cliente:
        flash("❌ Cliente no encontrado.", "danger")
        return redirect(url_for('inicio'))
    
    historial = gestor.obtener_historial(codigo)
    otros_prestamos = gestor.obtener_otros_prestamos(codigo, cliente['DNI'])
    
    cuotas_grid = []
    fecha_base = datetime.strptime(str(cliente['Fecha'])[:10], "%Y-%m-%d")
    cuotas_totales = int(cliente['Cuotas Total'])
    pagadas = int(cliente['Pagadas'])
    monto_cuota = float(cliente['Monto Cuota (S/)'])
    saldo_restante = float(cliente['Saldo (S/)'])
    
    for i in range(cuotas_totales):
        if cliente['Modalidad'] == "Diario": delta = timedelta(days=i+1)
        elif cliente['Modalidad'] == "Semanal": delta = timedelta(weeks=i+1)
        else: delta = timedelta(days=30*(i+1))
        
        monto_a_mostrar = monto_cuota
        if i == cuotas_totales - 1 and pagadas == i: 
            monto_a_mostrar = saldo_restante

        estado = "pagado" if i < pagadas else ("hoy" if i == pagadas else "pendiente")
        cuotas_grid.append({
            "numero": i + 1, "monto": monto_a_mostrar,
            "fecha": (fecha_base + delta).strftime("%d/%m/%Y"), "estado": estado
        })

    return render_template('tarjeta.html', cliente=cliente, historial=historial, cuotas_grid=cuotas_grid, otros_prestamos=otros_prestamos)

@app.route('/cobrar/<codigo>', methods=['POST'])
def cobrar(codigo):
    try:
        gestor.cobrar_cuota(codigo)
        flash("💰 Pago registrado en caja correctamente.", "success")
    except Exception as e:
        flash(f"⚠️ Error al procesar el pago: {e}", "danger")
    return redirect(url_for('detalle_cliente', codigo=codigo))

@app.route('/eliminar/<codigo>', methods=['POST'])
def eliminar(codigo):
    gestor.eliminar_cliente(codigo)
    flash("🗑️ Cliente y su historial de pagos eliminados por completo.", "info")
    return redirect(url_for('inicio'))

@app.route('/descargar-respaldo')
def descargar_respaldo():
    if not os.path.exists(ARCHIVO_PRINCIPAL):
        flash("⚠️ No hay datos para respaldar en este momento.", "warning")
        return redirect(url_for('inicio'))
        
    memory_file = io.BytesIO()
    
    with zipfile.ZipFile(memory_file, 'w') as zf:
        zf.write(ARCHIVO_PRINCIPAL, arcname=ARCHIVO_PRINCIPAL)
        if os.path.exists(ARCHIVO_HISTORIAL):
            zf.write(ARCHIVO_HISTORIAL, arcname=ARCHIVO_HISTORIAL)
            
        df = pd.read_excel(ARCHIVO_PRINCIPAL)
        fecha_hoy = datetime.now().strftime("%d/%m/%Y")
        
        texto_wsp = f"📱 MENSAJES PARA WHATSAPP - CREDIHOY ({fecha_hoy}) 📱\n"
        texto_wsp += "Copia el bloque correspondiente y envíalo a tu cliente:\n"
        texto_wsp += "="*50 + "\n\n"
        
        activos = df[df['Estado'] != 'Pagado']
        
        if activos.empty:
            texto_wsp += "No hay clientes con deuda pendiente para cobrar hoy.\n"
        else:
            for _, row in activos.iterrows():
                texto_wsp += f"Hola *{row['Nombres']}* 👋\n"
                texto_wsp += "Te compartimos el estado de tu préstamo con *CREDIHOY*:\n"
                texto_wsp += f"🔹 *Cuota {row['Modalidad']}:* S/ {float(row['Monto Cuota (S/)']):.2f}\n"
                texto_wsp += f"🔹 *Progreso:* {int(row['Pagadas'])} de {int(row['Cuotas Total'])} cuotas pagadas\n"
                texto_wsp += f"🔹 *Saldo Pendiente Total:* S/ {float(row['Saldo (S/)']):.2f}\n"
                texto_wsp += "¡Gracias por tu puntualidad! 🤝\n"
                texto_wsp += "-"*40 + "\n\n"
        
        zf.writestr("mensajes_whatsapp.txt", texto_wsp.encode('utf-8'))
        
    memory_file.seek(0)
    nombre_descarga = f"Respaldo_CREDIHOY_{datetime.now().strftime('%Y-%m-%d')}.zip"
    
    return send_file(memory_file, download_name=nombre_descarga, as_attachment=True, mimetype='application/zip')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
