"""RDA Pólizas D&A — controller.py: wires the GUI to the consolidation modules (confirm, run,
summarise in a popup, always re-enable the buttons in a `finally`).
Actions: execute_traspaso (step 1), execute_consolidacion (2), execute_poliza_sap (3: layout + batch),
execute_todo (2 and 3 in order). The processes share no state; 'todo' calls them in sequence and keeps score.
"""

import os
from tkinter import messagebox

from Comparativos import HOJA_COMPARATIVO_AGRUPADOR, HOJA_COMPARATIVO_BU
from Consolidacion import GRUPOS, HOJA_VARIACIONES, construir_variaciones, resolver_grupo
from Create_Poliza_Batch import generar_poliza_batch
from Create_Poliza_SAP import NOMBRES_HOJAS_LAYOUT, generar_poliza_sap
from Reservas_Manuales import previsualizar_pegado, traspasar_templates
from utils import claves_mes_actual_y_anterior, etiqueta_mes_es

# How many unmatched concepts to list inside the popup before trimming, so the message box never grows taller than the screen.
MAX_NO_MAPEADOS_EN_POPUP = 8


class ReservasController:
    """Connects the GUI with the two consolidations."""

    def __init__(self, gui):
        self.gui = gui
        self.gui.on_traspasar  = self.execute_traspaso
        self.gui.on_consolidar = self.execute_consolidacion
        self.gui.on_poliza     = self.execute_poliza_sap
        self.gui.on_todo       = self.execute_todo
        # The step 1 checklist lights a channel from the template FILE NAME (alphabetical order).
        self.gui.configurar_canales(sorted(g["canal"] for g in GRUPOS), self._canal_por_nombre)

    @staticmethod
    def _canal_por_nombre(ruta):
        """Channel a template FILE NAME points at ('ECOMM', 'HD'...), or None; only a hint for the GUI checklist."""
        indice = resolver_grupo(os.path.splitext(os.path.basename(ruta))[0])
        return None if indice is None else GRUPOS[indice]["canal"]

    # --- Shared runner ---
    def _ejecutar(self, funcion):
        """Run one consolidation -> (True, resumen) | (False, mensaje); touches no buttons or popups, so 'todo' survives a failed step."""
        try:
            return True, funcion()

        except (FileNotFoundError, ValueError) as e:
            # Expected, actionable problems: the message already says what to fix.
            return False, str(e)

        except PermissionError:
            return False, (
                "No pude escribir el resultado porque el archivo está abierto "
                "en Excel.\n\nCiérralo y vuelve a intentar."
            )

        except RuntimeError as e:
            # Raised by Poliza_SAP when pywin32 / Excel are not available.
            return False, str(e)

        except Exception as e:
            return False, f"Ocurrió un error inesperado:\n\n{e}"

    def _ejecutar_poliza(self, ruta_layout, ruta_plantilla):
        """Process 2 end to end: layout 'PÓLIZA SAP', then 'Póliza Carga Batch'. (True, resumen) when the layout worked
        (resumen['batch'] / ['batch_error'] carry the batch outcome), (False, mensaje) otherwise."""
        ok, resumen = self._ejecutar(lambda: generar_poliza_sap(
            ruta_layout, callback_status=self.gui.set_status,
        ))
        if not ok:
            return False, resumen

        self.gui.set_status("📦 Llenando la 'Póliza Carga Batch'...")
        # resumen['ruta'] is the processed copy in 'Output': the batch reads it and lands there too.
        ok_batch, payload_batch = self._ejecutar(lambda: generar_poliza_batch(
            resumen["ruta"], ruta_plantilla, callback_status=self.gui.set_status,
        ))
        resumen["batch"]       = payload_batch if ok_batch else None
        resumen["batch_error"] = None if ok_batch else payload_batch
        return True, resumen

    def _finalizar(self):
        """Re-enable the buttons and reset the status line if nothing succeeded."""
        self.gui.enable_buttons()
        if "completad" not in self.gui.status_var.get():
            self.gui.set_status("✓ Listo para comenzar")

    # --- PROCESO 1 — pestaña 'Variaciones' ---
    def _revisar_templates(self, ruta_libro, templates, mes):
        """Dry run of the template paste: validates everything, writes nothing, returns (ok, confirmation text | error).
        Runs BEFORE the confirmation so the person sees how many rows are pasted and how many 'Base' rows are replaced."""
        if not templates:
            return True, ""

        self.gui.set_status("🔎 Revisando los templates de Reservas Manuales...")
        ok, plan = self._ejecutar(
            lambda: previsualizar_pegado(ruta_libro, templates, mes)
        )
        if not ok:
            return False, plan
        return True, self._texto_confirmacion_pegado(plan)

    def execute_traspaso(self):
        """Paste the templates into 'Base' of a copy; the pivots are NOT built here."""
        if not self.gui.validate_traspaso():
            return

        ruta_libro = self.gui.get_archivo()
        mes        = self.gui.get_mes()
        templates  = self.gui.get_templates()

        # Buttons stay off for the whole thing, the review included, so a second click cannot start a second run.
        self.gui.disable_buttons()
        try:
            ok, texto_pegado = self._revisar_templates(ruta_libro, templates, mes)
            if not ok:
                self.gui.marcar_paso("traspaso", "error")
                self.gui.set_status("❌ Los templates no pasaron la validación")
                messagebox.showerror("No se puede traspasar", texto_pegado)
                return

            confirm = messagebox.askyesno(
                "Confirmar traspaso",
                f"¿Traspasar los templates a 'Base'?\n\n"
                f"  📄 Archivo : {os.path.basename(ruta_libro)}\n"
                f"{texto_pegado}\n"
                f"Se trabajará sobre una COPIA en la subcarpeta 'Output',\n"
                f"así que el archivo original no se toca."
            )
            if not confirm:
                return

            self.gui.marcar_paso("traspaso", "en_curso")
            ok, payload = self._ejecutar(lambda: traspasar_templates(
                ruta_libro, templates, mes=mes, callback_status=self.gui.set_status,
            ))
            if ok:
                self.gui.marcar_paso("traspaso", "listo")
                # The consolidation must run on the copy holding the pasted rows: point the file field at it.
                self.gui.set_archivo(payload["ruta"])
                self.gui.set_status("✅ ¡Traspaso completado!")
                messagebox.showinfo(
                    "Traspaso completado",
                    f"Archivo generado:\n{payload['ruta']}"
                    f"{self._texto_pegado(payload)}\n\n"
                    f"ℹ️ El campo del archivo ahora apunta a esta copia: el botón\n"
                    f"'Consolidar Variaciones' trabajará sobre ella.",
                )
            else:
                self.gui.marcar_paso("traspaso", "error")
                self.gui.set_status("❌ Error en el traspaso")
                messagebox.showerror("Error en el traspaso", payload)
        finally:
            self._finalizar()

    def execute_consolidacion(self):
        if not self.gui.validate_variaciones():
            return

        ruta_libro = self.gui.get_archivo()
        mes        = self.gui.get_mes()
        anterior, actual = claves_mes_actual_y_anterior(mes)

        confirm = messagebox.askyesno(
            "Confirmar consolidación",
            f"¿Generar la pestaña '{HOJA_VARIACIONES}'?\n\n"
            f"  📄 Archivo : {os.path.basename(ruta_libro)}\n"
            f"  📅 Compara : {anterior} vs {actual}\n"
            f"               ({etiqueta_mes_es(anterior)} vs {etiqueta_mes_es(actual)})\n"
            f"  📊 Además  : '{HOJA_COMPARATIVO_BU}' y\n"
            f"               '{HOJA_COMPARATIVO_AGRUPADOR}'\n"
            f"               (tablas dinámicas: año pasado arriba, año actual abajo)\n\n"
            f"Se trabajará sobre una COPIA en la subcarpeta 'Output',\n"
            f"así que el archivo original no se toca."
        )
        if not confirm:
            return

        self.gui.disable_buttons()
        try:
            self.gui.marcar_paso("consolidacion", "en_curso")
            ok, payload = self._ejecutar(lambda: construir_variaciones(
                ruta_libro, mes=mes, callback_status=self.gui.set_status,
            ))
            if ok:
                self.gui.marcar_paso("consolidacion", "listo")
                self.gui.set_status("✅ ¡Consolidación completada!")
                messagebox.showinfo(
                    "Consolidación completada",
                    self._texto_variaciones(payload),
                )
            else:
                self.gui.marcar_paso("consolidacion", "error")
                self.gui.set_status("❌ Error en la consolidación")
                messagebox.showerror("Error en la consolidación", payload)
        finally:
            self._finalizar()

    # --- PROCESO 2 — sección 'PÓLIZA SAP' ---
    def execute_poliza_sap(self):
        if not self.gui.validate_poliza():
            return

        ruta_libro     = self.gui.get_archivo_poliza()
        ruta_plantilla = self.gui.get_archivo_batch()

        confirm = messagebox.askyesno(
            "Confirmar póliza SAP",
            f"¿Llenar la sección 'PÓLIZA SAP' y la Póliza Carga Batch?\n\n"
            f"  📄 Layout  : {os.path.basename(ruta_libro)}\n"
            f"  📑 Pestañas: {' + '.join(NOMBRES_HOJAS_LAYOUT)}\n"
            f"  ✍️  Sección : 'PÓLIZA SAP' de cada pestaña\n"
            f"  📦 Plantilla: {os.path.basename(ruta_plantilla)}\n"
            f"     (se llena con los valores de la póliza)\n\n"
            f"Se abrirá Excel en segundo plano y se trabajará sobre COPIAS\n"
            f"en la subcarpeta 'Output'. Los originales no se tocan."
        )
        if not confirm:
            return

        self.gui.disable_buttons()
        try:
            self.gui.marcar_paso("poliza", "en_curso")
            ok, payload = self._ejecutar_poliza(ruta_libro, ruta_plantilla)
            if ok and not payload["batch_error"]:
                self.gui.marcar_paso("poliza", "listo")
                self.gui.set_status("✅ ¡Póliza SAP completada!")
                messagebox.showinfo(
                    "Póliza SAP completada", self._texto_poliza(payload)
                )
            elif ok:
                # Layout done and saved; only the batch step failed.
                self.gui.marcar_paso("poliza", "parcial")
                self.gui.set_status(
                    "⚠️ Póliza SAP completada, pero la Póliza Carga Batch falló"
                )
                messagebox.showwarning(
                    "Póliza SAP completada a medias", self._texto_poliza(payload)
                )
            else:
                self.gui.marcar_paso("poliza", "error")
                self.gui.set_status("❌ Error generando la póliza SAP")
                messagebox.showerror("Error generando la póliza SAP", payload)
        finally:
            self._finalizar()

    # --- Both, in order ---
    def execute_todo(self):
        if not self.gui.validate_todo():
            return

        ruta_var = self.gui.get_archivo()
        ruta_pol = self.gui.get_archivo_poliza()
        ruta_bat = self.gui.get_archivo_batch()
        mes      = self.gui.get_mes()
        anterior, actual = claves_mes_actual_y_anterior(mes)

        self.gui.disable_buttons()
        try:
            confirm = messagebox.askyesno(
                "Confirmar ejecución completa",
                f"¿Correr los pasos 2 y 3?\n\n"
                f"  2️⃣  '{HOJA_VARIACIONES}'\n"
                f"      📄 {os.path.basename(ruta_var)}\n"
                f"      📅 {anterior} vs {actual}\n\n"
                f"  3️⃣  'PÓLIZA SAP' + Póliza Carga Batch\n"
                f"      📄 {os.path.basename(ruta_pol)}\n"
                f"      📦 {os.path.basename(ruta_bat)}\n\n"
                f"Ambos escriben sobre COPIAS en la subcarpeta 'Output'."
            )
            if not confirm:
                return

            # --- Step 2 ---------------------------------------------------
            self.gui.marcar_paso("consolidacion", "en_curso")
            self.gui.set_status("2️⃣  Consolidando 'Variaciones'...")
            ok1, payload1 = self._ejecutar(lambda: construir_variaciones(
                ruta_var, mes=mes, callback_status=self.gui.set_status,
            ))
            self.gui.marcar_paso("consolidacion", "listo" if ok1 else "error")

            if not ok1:
                seguir = messagebox.askyesno(
                    "Falló el paso 2",
                    f"La consolidación de '{HOJA_VARIACIONES}' falló:\n\n"
                    f"{payload1}\n\n"
                    f"¿Continúo con el paso 3 ('PÓLIZA SAP')?",
                )
                if not seguir:
                    self.gui.set_status("❌ Ejecución detenida en el paso 2")
                    return

            # --- Step 3 ---------------------------------------------------
            self.gui.marcar_paso("poliza", "en_curso")
            self.gui.set_status("3️⃣  Generando la 'PÓLIZA SAP'...")
            ok2, payload2 = self._ejecutar_poliza(ruta_pol, ruta_bat)
            # Step 3 counts as fully OK only if the batch file was also produced.
            batch_fallo = ok2 and bool(payload2["batch_error"])
            self.gui.marcar_paso(
                "poliza", "error" if not ok2 else ("parcial" if batch_fallo else "listo"))

            # --- Combined report ------------------------------------------
            partes = []
            partes.append(
                f"2️⃣  {HOJA_VARIACIONES}  —  {'✅ OK' if ok1 else '❌ FALLÓ'}\n"
                + (self._texto_variaciones(payload1) if ok1 else payload1)
            )
            estado2 = "❌ FALLÓ" if not ok2 else ("⚠️ A MEDIAS" if batch_fallo else "✅ OK")
            partes.append(
                f"3️⃣  PÓLIZA SAP + Póliza Carga Batch  —  {estado2}\n"
                + (self._texto_poliza(payload2) if ok2 else payload2)
            )
            texto = ("\n" + "─" * 48 + "\n").join(partes)

            if ok1 and ok2 and not batch_fallo:
                self.gui.set_status("✅ ¡Ejecución completa completada!")
                messagebox.showinfo("Ejecución completa", texto)
            elif ok1 or ok2:
                self.gui.set_status("⚠️ Ejecución completada a medias")
                messagebox.showwarning("Ejecución parcial", texto)
            else:
                self.gui.set_status("❌ Fallaron los dos procesos")
                messagebox.showerror("Ejecución fallida", texto)
        finally:
            self._finalizar()

    # --- Summary popups ---
    @staticmethod
    def _texto_variaciones(resumen):
        """Build the 'Variaciones' popup text."""
        bloques = "\n".join(
            f"     • {titulo}: {n} concepto(s)"
            for titulo, n in resumen["conceptos_por_bloque"].items()
        )

        texto = (
            f"Archivo generado:\n{resumen['ruta']}\n\n"
            f"  • Meses comparados : {resumen['mes_anterior']} vs {resumen['mes_actual']}\n"
            f"  • Filas leídas     : {resumen['filas_base']:,}\n"
            f"  • Matrices         : {resumen['bloques']}\n"
            f"{bloques}\n"
            f"  • Σ {resumen['mes_anterior']}        : {resumen['total_anterior']:,.2f}\n"
            f"  • Σ {resumen['mes_actual']}        : {resumen['total_actual']:,.2f}"
        )

        texto += ReservasController._texto_pegado(resumen.get("pegado"))
        texto += ReservasController._texto_comparativos(resumen)

        no_mapeados = resumen.get("conceptos_no_mapeados") or []
        if no_mapeados:
            mostrados = no_mapeados[:MAX_NO_MAPEADOS_EN_POPUP]
            resto = len(no_mapeados) - len(mostrados)
            texto += (
                f"\n\n⚠️ {len(no_mapeados)} valor(es) de 'Agrupador' no empataron "
                f"con el catálogo de conceptos.\n"
                f"Salieron como renglón propio al final de su matriz, con su "
                f"nombre original (los totales sí cuadran):\n"
                + "\n".join(f"     • {c}" for c in mostrados)
                + (f"\n     • ...y {resto} más" if resto > 0 else "")
            )

        sin_bloque = resumen.get("canales_sin_bloque") or []
        if sin_bloque:
            mostrados = sin_bloque[:MAX_NO_MAPEADOS_EN_POPUP]
            resto = len(sin_bloque) - len(mostrados)
            texto += (
                f"\n\n⚠️ Canal(es) de la Base sin matriz asignada "
                f"({resumen['mes_anterior']} y {resumen['mes_actual']}):\n"
                + "\n".join(f"     • {c}" for c in mostrados)
                + (f"\n     • ...y {resto} más" if resto > 0 else "")
                + "\nSus importes NO aparecen en ninguna matriz."
            )

        return texto

    @staticmethod
    def _desglose_hojas(hojas):
        """'Manuales 2 + Rentas 1': rows taken from each tab of a template."""
        return " + ".join(f"{h['etiqueta']} {h['filas']}" for h in hojas)

    @staticmethod
    def _texto_confirmacion_pegado(plan):
        """Template half of the confirmation box: what will be pasted and what it will REPLACE (from a dry-run plan)."""
        lineas = "\n".join(
            f"     • {a['canal']:<11} {len(a['filas']):>4} fila(s)  Σ {a['total']:>16,.2f}"
            f"   [{ReservasController._desglose_hojas(a['hojas'])}]   ({a['nombre']})"
            for a in plan["archivos"]
        )
        texto = (
            f"\n  📥 Se pegarán en 'Base' ({plan['mes']}):\n{lineas}\n"
            f"     Total: {len(plan['filas_nuevas']):,} fila(s)  Σ {plan['total_nuevo']:,.2f}\n"
        )

        if plan["filas_a_reemplazar"]:
            por_canal = ", ".join(
                f"{canal} {n}" for canal, n in plan["reemplazo_por_canal"].items() if n
            )
            texto += (
                f"\n  ♻️  Se REEMPLAZARÁN {len(plan['filas_a_reemplazar']):,} fila(s) que ya "
                f"están en 'Base'\n"
                f"     para {plan['mes']} y esos mismos canales (Σ {plan['total_reemplazado']:,.2f}):\n"
                f"     {por_canal}\n"
            )
        else:
            texto += (
                f"\n  ✅ 'Base' todavía no tiene {plan['mes']} de estos canales: "
                f"no se reemplaza nada.\n"
            )

        for aviso in plan["advertencias"]:
            texto += f"  ⚠️ {aviso}\n"
        return texto

    @staticmethod
    def _texto_pegado(pegado):
        """The 'templates pasted' section of the result popup ('' if none)."""
        if not pegado:
            return ""

        lineas = "\n".join(
            f"     • {a['canal']}: {a['filas']:,} fila(s), Σ {a['total']:,.2f}"
            f"   [{ReservasController._desglose_hojas(a['hojas'])}]"
            for a in pegado["archivos"]
        )
        texto = (
            f"\n\n📥 Templates pegados en 'Base' ({pegado['mes']}):\n{lineas}\n"
            f"  • Filas pegadas     : {pegado['filas_pegadas']:,} "
            f"(filas {pegado['fila_inicio']}–{pegado['fila_fin']}), "
            f"Σ {pegado['total_pegado']:,.2f}\n"
            f"  • Filas reemplazadas: {pegado['filas_reemplazadas']:,}"
        )
        if pegado["filas_reemplazadas"]:
            texto += f" (Σ {pegado['total_reemplazado']:,.2f})"
        if pegado.get("verificado"):
            texto += "\n  ✅ Verificado: lo pegado coincide con los templates, celda por celda."
        for aviso in pegado["advertencias"]:
            texto += f"\n  ⚠️ {aviso}"
        return texto

    @staticmethod
    def _texto_comparativos(resumen):
        """PivotTable half of the 'Variaciones' popup: built, skipped on purpose, or failed.
        A failure is not an error dialog (the sheet is already saved) but must not be silent either."""
        error = resumen.get("comparativos_error")
        if error:
            return (
                f"\n\n⚠️ Las tablas dinámicas comparativas NO se crearon:\n"
                f"{error}\n"
                f"La pestaña 'Variaciones' sí quedó lista en el archivo de arriba.\n"
                f"(Ese paso necesita Excel instalado y el archivo cerrado.)"
            )

        comparativos = resumen.get("comparativos")
        if not comparativos:
            return ""

        lineas = []
        for tabla in comparativos["tablas"]:
            if tabla["tabla"] is None:
                lineas.append(
                    f"     • {tabla['hoja']} — {tabla['anio']}: sin movimientos"
                )
            else:
                meses = tabla["meses"]
                total = tabla["total"]
                lineas.append(
                    f"     • {tabla['hoja']} — {tabla['anio']}: "
                    f"{len(meses)} mes(es) ({meses[0]}–{meses[-1]})"
                    + (f", Σ {total:,.2f}" if total is not None else "")
                )

        texto = (
            f"\n\n📊 Tablas dinámicas creadas "
            f"({comparativos['anio_anterior']} arriba / "
            f"{comparativos['anio_actual']} abajo):\n"
            + "\n".join(lineas)
        )

        for aviso in comparativos.get("advertencias") or []:
            texto += f"\n     ⚠️ {aviso}"

        return texto

    @staticmethod
    def _texto_poliza(resumen):
        """Build the 'PÓLIZA SAP' popup text. The #N/A warning is the important half: those rows go up to SAP,
        so an unknown account, cost centre or description in 'Catalogo' must be fixed BEFORE the upload."""
        columnas = "\n".join(f"     • {c}" for c in resumen["columnas"])
        hojas = "\n".join(
            f"     • {h['hoja']}: {h['filas']:,} renglón(es) "
            f"(filas {h['fila_inicio']}–{h['fila_fin']}), Σ {h['total_control']:,.2f}"
            for h in resumen["hojas"]
        )

        texto = (
            f"Archivo generado:\n{resumen['ruta']}\n\n"
            f"  • Pestañas         :\n{hojas}\n"
            f"  • Renglones        : {resumen['filas']:,}\n"
            f"  • Columnas         :\n{columnas}\n"
            f"  • Σ {resumen['columna_control']} (todas las pestañas) : "
            f"{resumen['total_control']:,.2f}\n"
        )

        errores = resumen.get("errores") or []
        if errores:
            texto += (
                f"\n\n⚠️ {resumen['total_errores']} celda(s) con error en los "
                f"VLOOKUP — hay que corregirlas ANTES de subir a SAP:\n"
            )
            for e in errores:
                celdas = ", ".join(f"{c['celda']} {c['error']}" for c in e["celdas"])
                resto = e["total"] - len(e["celdas"])
                texto += (
                    f"     • {e['hoja']} — {e['columna']} ({e['titulo']}): "
                    f"{e['total']} error(es)\n"
                    f"       {celdas}"
                    + (f", ...y {resto} más" if resto > 0 else "")
                    + "\n"
                )
            texto += (
                "\nNormalmente significa que el valor no existe en la pestaña "
                "'Catalogo'."
            )
        else:
            texto += "\n\n✅ Sin errores: todos los VLOOKUP empataron."

        avisos = resumen.get("advertencias") or []
        if avisos:
            texto += "\n\nℹ️ Avisos:\n" + "\n".join(f"     • {a}" for a in avisos)

        texto += ReservasController._texto_batch(resumen)

        return texto

    @staticmethod
    def _texto_batch(resumen):
        """'Póliza Carga Batch' half of the popup: produced (with its warnings) or not.
        Not producing it is no error dialog (the layout is already saved) but must not be silent."""
        error = resumen.get("batch_error")
        if error:
            return (
                f"\n\n❌ La Póliza Carga Batch NO se generó:\n\n{error}\n\n"
                f"La 'PÓLIZA SAP' del Layout sí quedó lista en el archivo de arriba."
            )

        batch = resumen.get("batch")
        if not batch:
            return ""

        d_ini, d_fin = batch["bloque_debito"]
        c_ini, c_fin = batch["bloque_credito"]
        origen = " + ".join(f"{h['hoja']} {h['filas']}" for h in batch["hojas_origen"])
        texto = (
            f"\n\n📦 Póliza Carga Batch generada:\n{batch['ruta']}\n\n"
            f"  • Pestaña          : {batch['hoja']}\n"
            f"  • Origen (Layout)  : {origen}\n"
            f"  • Bloque débito    : filas {d_ini}–{d_fin} ({batch['filas_origen']:,})\n"
            f"  • Bloque crédito   : filas {c_ini}–{c_fin} ({batch['filas_origen']:,})\n"
            f"  • Σ Debit Amount   : {batch['total_debito']:,.2f}\n"
            f"  • Σ Credit Amount  : {batch['total_credito']:,.2f}\n"
            f"  ✅ Verificado: lo pegado coincide con el Layout, celda por celda."
        )

        advertencias = batch.get("advertencias") or []
        if advertencias:
            texto += "\n\n⚠️ Revisa antes de subir a SAP:\n" + "\n".join(
                f"     • {a}" for a in advertencias
            )

        texto += (
            "\n\nℹ️ Todavía sin llenar en la plantilla: 'GL Account' del bloque de\n"
            "crédito y los campos naranjas (Company Code, fechas, etc.)."
        )
        return texto
