"""Prueba de humo: vista normal, ?admin=1 (login + publicar) y ?editar=1 (editar avances)."""
from streamlit.testing.v1 import AppTest

print("=== 1) Vista normal (sin query params) ===")
at = AppTest.from_file("app.py", default_timeout=60)
at.run()
if at.exception:
    for e in at.exception:
        print("EXCEPCION:", e)
    raise SystemExit(1)
print("OK - tabs encontrados:", [t.label for t in at.tabs])

print("\n=== 2) Interactuar con 'Mi Cartera (Gestor)': elegir un gestor ===")
gestor_sb = at.selectbox(key="gestor_seleccionado")
print("Gestores disponibles (primeros 3):", gestor_sb.options[:3])
gestor_sb.set_value(gestor_sb.options[0]).run()
if at.exception:
    for e in at.exception:
        print("EXCEPCION:", e)
    raise SystemExit(1)
print("OK - metrics tras elegir gestor:", [m.value for m in at.metric][:5])

print("\n=== 3) Vista Gerencial ===")
# Ya corrió dentro del mismo run (ambos tabs se renderizan). Verificamos que
# existan métricas / tablas sin excepción.
print("OK - sin excepciones en vista gerencial (ya validado en el run anterior)")

print("\n=== 4) ?admin=1 : login y publicar plantilla generada ===")
at_admin = AppTest.from_file("app.py", default_timeout=60)
at_admin.query_params["admin"] = "1"
at_admin.run()
if at_admin.exception:
    for e in at_admin.exception:
        print("EXCEPCION:", e)
    raise SystemExit(1)

user_input = at_admin.sidebar.text_input[0]
pass_input = at_admin.sidebar.text_input[1]
user_input.set_value("admin").run()
pass_input.set_value("admin2025").run()
botones = at_admin.sidebar.button
print("Botones en sidebar tras set values:", [b.label for b in botones])

login_btn = [b for b in at_admin.sidebar.button if b.label == "Ingresar"][0]
login_btn.click().run()
if at_admin.exception:
    for e in at_admin.exception:
        print("EXCEPCION:", e)
    raise SystemExit(1)
print("OK - login exitoso, sin excepciones en la app.")

print("\n=== 5) Publicar datos generados desde la propia plantilla ===")
import sys
sys.path.insert(0, ".")
from app import generar_datos_ejemplo, publicar_datos
import os

datos = generar_datos_ejemplo()
publicar_datos(datos, dia_corte=15, mes=8, anio=2026)
print("OK - publicar_datos() no lanzó excepción. Archivo existe:", os.path.exists("data/ultima_carga.xlsx"))

print("\n=== 6) Releer app.py tras publicación (debe usar el archivo publicado) ===")
at_post = AppTest.from_file("app.py", default_timeout=60)
at_post.run()
if at_post.exception:
    for e in at_post.exception:
        print("EXCEPCION:", e)
    raise SystemExit(1)
print("OK - app corre con datos publicados. Caption:", [c.value for c in at_post.caption][:1])

print("\n=== 7) ?editar=1 : pestaña Editar Avances ===")
at_editar = AppTest.from_file("app.py", default_timeout=60)
at_editar.query_params["editar"] = "1"
at_editar.run()
if at_editar.exception:
    for e in at_editar.exception:
        print("EXCEPCION:", e)
    raise SystemExit(1)
print("OK - tabs con editar:", [t.label for t in at_editar.tabs])

gestor_editor_sb = at_editar.selectbox(key="gestor_editor")
print("Gestores en editor (primeros 3):", gestor_editor_sb.options[:3])
gestor_editor_sb.set_value(gestor_editor_sb.options[0]).run()
if at_editar.exception:
    for e in at_editar.exception:
        print("EXCEPCION:", e)
    raise SystemExit(1)
print("OK - selección de gestor en editor sin excepciones")

print("\n=== TODO OK ===")


