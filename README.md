# Herramienta SIRCE · Capacitación

Arma el layout de carga de capacitaciones para SIRCE (STPS) a partir de los reportes en Excel.
Todo se procesa en el navegador de quien la usa: **los archivos no se suben a ningún servidor**.

## Uso
1. **Archivos:** reportes de capacitación (NOM), catálogo de colaboradores, catálogo de claves, claves de cursos, información de sucursales (clave de establecimiento, razón social, RFC y representante de los trabajadores) y, opcional, cargas anteriores a SIRCE.
2. **Revisión:** corrige fechas dudosas, puestos, escolaridades, sucursales y cursos sin clave. Excluye cursos o compañías que no apliquen.
3. **Equivalencias:** alias de empresa, municipio principal por estado, y exportar/importar equivalencias para compartirlas con el equipo.
4. **Generar:** descarga el Excel con una pestaña por empresa (27 columnas).
5. **DC-3:** PDF por colaborador y curso (anverso y reverso) en ZIP, con una carpeta por sucursal. Nombre: SUCURSAL - COLABORADOR - CURSO - DC3.

## Reglas
- Fechas del reporte NOM con día/mes invertidos se corrigen con la columna *Mes*.
- Quien no está en el catálogo se considera baja y se ignora.
- Municipio = municipio principal del estado; las claves salen de la hoja *Claves*.
- Clave de curso = nombre del curso + fecha de inicio + empresa (archivo de claves de cursos). Área temática siempre 3134. Duración = horas del reporte.

## Privacidad
No subir al repositorio archivos con CURP, catálogos ni la hoja de claves (contiene usuarios y contraseñas de SIRCE). El `.gitignore` bloquea `.xls`, `.xlsx`, `.csv`, `.pdf` y `.docx`.

## Publicar (GitHub Pages)
Repositorio → Settings → Pages → *Deploy from a branch* → `main` / `(root)`.

## Pendiente
- DC-2 (plan y programa de capacitación) por sucursal con catálogo de temarios editable.
