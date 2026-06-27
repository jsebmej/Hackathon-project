DataLex 1581 — Interfaz web
============================

Autodiagnóstico inteligente para la protección de datos (Ley 1581 de 2012).

ESTRUCTURA DE ARCHIVOS
----------------------
datalex-1581/
├── index.html        → Pantalla 1: Inicio (landing)
├── styles.css            estilos del inicio
├── login.html        → Pantalla 2: Registro / Inicio de sesión
├── login.css             estilos del registro
├── diagnostico.html  → Pantalla 3: Panel de diagnóstico (chat DataLex AI)
├── diagnostico.css       estilos del diagnóstico
└── README.txt        → este archivo

FLUJO DE NAVEGACIÓN
-------------------
   index.html
       │  (botón "Regístrate aquí")
       ▼
   login.html
       │  (botón "Ingresar")
       ▼
   diagnostico.html

   · El logo de cada pantalla regresa al inicio (index.html).
   · El nombre de empresa escrito en el registro se conserva como contexto del diagnóstico.

CÓMO USARLA
-----------
1. Mantén los 6 archivos dentro de la misma carpeta (no los separes; los
   enlaces son relativos).
2. Abre "index.html" con cualquier navegador (doble clic) y recorre el flujo.

NOTAS
-----
· La tipografía Inter se carga desde Google Fonts (requiere conexión la
  primera vez; si no, usa una fuente del sistema como respaldo).
· El botón "Ingresar" aún no valida campos ni se conecta a un backend:
  redirige directamente al diagnóstico. El chat es un mock interactivo
  (tus respuestas aparecen como burbujas), listo para conectar tu motor de IA.
· Paleta: azul oscuro #0B1E4A–#15306B, azul #2563EB, verde #22C55E/#16A34A,
  grises #374151/#9CA3AF, fondos suaves #F4F7FC y blanco #FFFFFF.

Si prefieres todo en un solo archivo autónomo, usa "datalex.html"
(versión de una sola página con las tres vistas integradas).
