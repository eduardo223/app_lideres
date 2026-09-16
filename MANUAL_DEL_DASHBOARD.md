# 📘 MANUAL MAESTRO DE OPERACIÓN Y GESTIÓN ESTRATÉGICA — DASHBOARD LÍDERES & GERENCIA

Bienvenida/o a la guía definitiva de control, análisis y toma de decisiones comerciales de la plataforma. Este manual fue diseñado para transformar la visualización de datos en **estrategia de alto nivel**, permitiéndote sustentar informes gerenciales con solvencia técnica y liderar tu sector con precisión numérica.

---

## 🧭 TABLA DE CONTENIDOS
1. [¿De dónde sale la tarjeta de Activas (Real / Obj)? (Caso Screenshot)](#1-de-dónde-sale-la-tarjeta-de-activas-real--obj-caso-screenshot)
2. [Matriz de Fuentes de Datos: ¿Cómo se alimenta el Dashboard?](#2-matriz-de-fuentes-de-datos-cómo-se-alimenta-el-dashboard)
3. [Guía de Carga Paso a Paso desde la Barra Lateral](#3-guía-de-carga-paso-a-paso-desde-la-barra-lateral)
4. [Diccionario Ejecutivo de KPIs Superiores](#4-diccionario-ejecutivo-de-kpis-superiores)
5. [Pirámide de Niveles (Tier Cards: Bronce a Diamante)](#5-pirámide-de-niveles-tier-cards-bronce-a-diamante)
6. [Módulos y Pestañas de Navegación Analítica](#6-módulos-y-pestañas-de-navegación-analítica)
7. [Cómo Presentar Informes Comerciales con Alto Nivel](#7-cómo-presentar-informes-comerciales-con-alto-nivel)
8. [Rutina Estratégica por Momentos del Ciclo](#8-rutina-estratégica-por-momentos-del-ciclo)

---

## 1. ¿De dónde sale la tarjeta de Activas (Real / Obj)? (Caso Screenshot)

![Tarjeta Activas](attachment:captura)

### 📌 Ficha Técnica del Indicador:
* **Etiqueta en Pantalla:** `👥 ACTIVAS (REAL / OBJ)`
* **Valor Principal (`807`):** Total acumulado de consultoras del sector (o del grupo de la líder) que han pasado pedido con corte oficial en el ciclo activo.
* **Badge / Delta (`↑ 83.1% Cumplimiento`):** Porcentaje de avance frente al objetivo formal fijado por la compañía.
  $$\% \text{ Cumplimiento} = \left( \frac{\text{Real Activas}}{\text{Objetivo Activas}} \right) \times 100$$
  *(Para 807 con 83.1%, la meta de tu sector es de aproximadamente 971 activas).*
* **¿Qué es la lupa `🔍` y el botón?** Al hacer clic sobre la tarjeta o su botón de auditoría, se abre el diálogo interactivo **`dialog_origen_activas`**, que te muestra el desglose exacto de cuántas activas aporta cada Líder de Negocio de tu sector.
* **¿Por qué sale abajo el texto *"Pica aquí para ver análisis y opciones de PLATA"*?** Ese mensaje corresponde al **tooltip de la tarjeta de nivel PLATA**, que está posicionada inmediatamente debajo del cuadro de mando. Al pasar el cursor sobre ella, te invita a desplegar el diagnóstico profundo de consultoras de dicho nivel.

### 📥 Archivo y Columna que lo alimenta:
* **Archivo Fuente:** `Base para el como vamos.xlsx`
* **Hoja Requerida:** `Base para el como vamos`
* **Columnas exactas de cálculo:**
  1. `Real Activas`: Conteo oficial acumulado de pedidos de ciclo.
  2. `Objetivo Activas`: Meta fijada por Natura para el ciclo.
  3. `Código de grupo` / `Nombre de consultora`: Identificador de la líder.
  4. `Nombre Setor` / `Sector`: Filtro de seguridad multitenant de tu sector.

---

## 2. Matriz de Fuentes de Datos: ¿Cómo se alimenta el Dashboard?

El sistema unifica 6 fuentes de información comercial y financiera. Cada archivo alimenta un área especializada del cuadro de mando:

| # | Archivo Oficial | Origen en Natura | Módulos que Alimenta | Frecuencia Sugerida |
|---|---|---|---|---|
| **1** | **`Base de Datos.xlsx`** (Tableau) | Portal Tableau Natura (Reporte de Consultoras) | Directorio 1 a 1 de consultoras, pestaña *Informe Tableau Cam*, datos de contacto, cumpleaños, niveles de carrera. | Inicio de campaña y semanal. |
| **2** | **`mi_grupo.xls`** o **`activas`** | Portal Mi Negocio / Reporte operativo diario | Pestaña *Tableau*: actualiza en caliente quién ya facturó hoy sin esperar a descargar toda la base pesada. | Diario (en los días de alta facturación). |
| **3** | **`Base para el como vamos.xlsx`** | Reporte Comercial "Cómo Vamos" | **Cuadro de mando superior**: Activas Reales, Facturación Real, Metas, Brechas al 100%/110%, Tacómetros 360°, Ganancia estimada LN. | Cada 2 a 3 días durante el ciclo. |
| **4** | **`Gera.xlsx` / `Geral.xlsx`** | Módulo de Crédito y Cobranza Gera | Pestaña *Gera_Credito&Cobranza*: Títulos vencidos, saldo en mora, riesgo crediticio, seguimiento de facturas impagas. | 2 veces por semana (clave antes de cierre). |
| **5** | **`Objetivos Arte.xlsx`** | Hoja oficial *Desafíos LNN* | Metas de **Inicios, Reinicios y Recuperos** de cada líder y consolidado de sector. | Al inicio del ciclo (una sola vez). |
| **6** | **`Ajustes Desafíos`** (Excel Zona) | Plantilla de Gerencia de Zona | Conciliación de metas calibradas por zona vs Objetivos Arte corporativos. | Inicio de ciclo si hubo negociación de metas. |

---

## 3. Guía de Carga Paso a Paso desde la Barra Lateral

En el menú lateral izquierdo encontrarás el panel **"Carga y Actualización de Bases"**, organizado en pestañas numeradas para evitar confusiones:

```
[ Barra Lateral ]
  ├── 📊 1. Tableau (Base Maestra de Consultoras 1 a 1)
  │     ├── 📁 Tableau (Base de Datos.xlsx)
  │     ├── 🔄 mi_grupo (Cruces rápidos del día)
  │     └── ⚡ Activas (Actualización por lote de pedidos)
  ├── 🔄 2. Cómo Vamos (Metas de Ciclo y Facturación)
  ├── 💳 3. Gera (Crédito y Cartera Vencida)
  ├── 🎯 4. Objetivos Arte (Inicios / Reinicios / Recuperos corporativos)
  ├── ✨ 5. Desafíos (Calibración de Gerencia de Zona)
  ├── 📑 6. Ganancia Arte (CE+ y Comisiones Especiales)
  └── 📝 7. Historial (Trazabilidad y logs de auditoría)
```

### Protocolo para cargar un nuevo ciclo:
1. **Paso 1 (Cómo Vamos):** Entra a la pestaña **`🔄 2. Cómo Vamos`**, sube el archivo descargado de Natura y presiona **"🚀 Rotar Ciclo y Actualizar Histórico"**. El sistema guardará el ciclo previo como comparativo histórico y colocará el nuevo en el cuadro de mando.
2. **Paso 2 (Tableau):** Entra a **`📊 1. Tableau`**, sube la `Base de Datos.xlsx`. El sistema sincronizará la base de datos SQLite y detectará automáticamente si ingresaron nuevas líderes de negocio para habilitarles su cuenta.
3. **Paso 3 (Gera):** Entra a **`💳 3. Gera`**, sube la exportación de Gera (en formato Excel Inmediata) para tener el mapa de riesgo crediticio actualizado.
4. **Paso 4 (Objetivos Arte):** Entra a **`🎯 4. Objetivos Arte`** para actualizar las metas de nuevos ingresos y recuperación de red.

---

## 4. Diccionario Ejecutivo de KPIs Superiores

Cuando presentes ante tu equipo o gerencia regional, utiliza estas definiciones oficiales:

### 1. `👥 CONSULTORAS / LÍDERES`
* **Definición:** Número total de grupos comerciales formalmente estructurados dentro de la gerencia de sector.
* **Uso Gerencial:** Dimensión operativa de tu equipo directo.

### 2. `👥 ACTIVAS (REAL / OBJ)`
* **Definición:** Consultoras con pedido procesado frente a la meta mínima exigida.
* **Lectura:** Permite identificar si el sector alcanzará la cuota de actividad antes de que se agoten los días del ciclo.

### 3. `💰 FACTURACIÓN TOTAL`
* **Definición:** Venta neta comercial acumulada en millones de pesos ($M COP) y su porcentaje frente a la meta.
* **Cálculo Clave:** Brecha de Facturación:
  $$\text{Brecha al 100\%} = \text{Objetivo Facturación} - \text{Real Facturación}$$
  $$\text{Brecha al 110\% (Superación)} = (\text{Objetivo Facturación} \times 1.10) - \text{Real Facturación}$$

### 4. `💵 GANANCIA ESTIMADA LN`
* **Definición:** Estimación del ingreso que percibirán las Líderes de Negocio según la matriz de comisiones de Natura y la aplicación de los factores multiplicadores (potencializador de saldo comercial positivo).

### 5. `⚖️ SALDO COMERCIAL`
* **Definición:** Balance neto del crecimiento de la red:
  $$\text{Saldo Comercial} = (\text{Inicios} + \text{Reinicios} + \text{Recuperos}) - \text{Cesantes (Fugas a I4)}$$
* **Meta Estándar:** $+2$ por grupo o sector. Si el saldo es positivo, la red está expandiéndose; si es negativo, la red se está descapitalizando.

### 6. `🚀 INICIOS / REINICIOS`
* **Definición:** Nuevas consultoras que debutan en la campaña (Inicios) más consultoras que habían cesado y volvieron a ingresar formalmente (Reinicios).
* **Impacto:** Motor principal de reposición de red y crecimiento futuro.

### 7. `🌸 BOLSA DE RECUPERACIÓN DE RED`
* **Inactiva 1 (I1):** 1 ciclo sin pasar pedido. Tasa de conversión óptima (campaña suave de reactivación).
* **Inactiva 2 (I2):** 2 ciclos sin pasar pedido. Alerta de desmotivación o problema de crédito.
* **Inactiva 3 (I3) — ¡Riesgo Fuga a I4!:** 3 ciclos sin pedido. Última oportunidad comercial antes de que la consultora sea dada de baja como cesante del sistema.
* **Recuperos Logrados:** Consultoras rescatadas de la bolsa I1-I3 que ya pasaron pedido en el ciclo actual.

---

## 5. Pirámide de Niveles (Tier Cards: Bronce a Diamante)

En el bloque inferior del cuadro de mando encontrarás 5 tarjetas con la identidad de color de Natura:
* 🥉 **BRONCE** (Base del negocio)
* 🥈 **PLATA** (Consolidación)
* 🥇 **ORO** (Desarrollo y volumen)
* 💎 **ZAFIRO** (Alto impacto y liderazgo)
* 👑 **DIAMANTE** (Cúspide de facturación y fidelidad)

### ¿Cómo operar la consola de Deep Dive por Nivel?
1. Al presionar el botón con la lupa `🔍` en cualquiera de las tarjetas (por ejemplo, **PLATA**), la plataforma activa el **Centro de Diagnóstico Dinámico**.
2. Te mostrará:
   * Total de consultoras de ese nivel.
   * % de Actividad específica del nivel.
   * Ticket promedio de facturación ($ COP por consultora activa de ese nivel).
   * Listado directo con números de teléfono y enlace directo a WhatsApp para lanzar campañas de incentivo por segmento.
3. Para salir de ese filtro y volver al panorama global, simplemente vuelve a pulsar sobre el botón de la tarjeta (ahora marcado con una `✖️`).

---

## 6. Módulos y Pestañas de Navegación Analítica

### 📊 Pestaña 1: Informe Tableau Cam / Mi Listado
* **Finalidad:** Gestión 1 a 1 de las consultoras de la red.
* **Herramientas Clave:**
  * **Filtros Combinados:** Por Líder, Nivel, Situación Comercial (Activa, I1, I2, I3, Cesante), Estado de Pedido (Paso pedido / Sin pedido).
  * **Acción WhatsApp Directa:** Envío de mensajes personalizados con un solo clic con plantilla de saludo y recordatorio de ciclo.
  * **Monitor de Cumpleaños:** Banner superior para felicitar a consultoras en su fecha especial como herramienta de relacionamiento.

### 🔄 Pestaña 2: Cómo Vamos / Tacómetros 360°
* **Finalidad:** Seguimiento de metas financieras y comerciales.
* **Herramientas Clave:**
  * **Tacómetros de Velocidad:** Visualizadores circulares estilo velocímetro con zonas de color (Rojo: <95%, Amarillo: 95-99%, Verde: >=100%).
  * **Tabla de Ranking de Desempeño:** Ordenada de mayor a menor cumplimiento de facturación, calculando la brecha en pesos COP para el 100%.

### 💳 Pestaña 3: Gera Crédito & Cobranza
* **Finalidad:** Proteger la rentabilidad y asegurar el cobro oportuno.
* **Herramientas Clave:**
  * Semáforo de Cartera Vencida por rangos de días (1 a 30 días, 31 a 60 días, 61+ días).
  * Monto total expuesto a riesgo crediticio.
  * Títulos de deuda descargables para gestionar acuerdos de pago antes del corte de facturación de nuevos pedidos.

### 🎯 Pestaña 4: Objetivos Arte & Desafíos
* **Finalidad:** Planificación estratégica y auditoría de metas.
* **Herramientas Clave:**
  * Tablero de conciliación: Compara lo que exige Natura corporativo vs los ajustes que la Gerente de Zona pactó con cada líder.
  * Histórico de campañas para evaluar la tendencia de cumplimiento ciclo a ciclo.

---

## 7. Cómo Presentar Informes Comerciales con Alto Nivel

Para transformar una reunión operativa en una presentación directiva de alto impacto, adopta estas buenas prácticas de comunicación cuantitativa:

| En vez de decir... ❌ | Expresa con nivel gerencial... ✅ |
|---|---|
| *"Nos faltan unas cuantas consultoras para llegar"* | *"Nuestra tasa de actividad actual está en el 83.1% con 807 activas; la brecha crítica para el 100% es de 164 pedidos"* |
| *"La gente no está comprando mucho"* | *"El ticket promedio se sitúa en $X COP; observamos una contracción del 12% en el segmento Bronce que debemos compensar con volumen en Plata y Oro"* |
| *"Hay muchas consultoras que no pasan pedido"* | *"Tenemos 42 consultoras en riesgo I3; si no recuperamos al menos el 30% antes del cierre, impactará negativamente nuestro Saldo Comercial en -13 puntos netos"* |
| *"Las líderes van bien"* | *"El 70% del equipo ya superó el umbral del 95% de cumplimiento de facturación; el foco de acompañamiento en terreno se concentra en los 3 grupos con mayor brecha al 100%"* |

---

## 8. Rutina Estratégica por Momentos del Ciclo

Un/a líder o gerente de alto desempeño gestiona el dashboard en 3 fases:

### 🟢 Fase 1: Apertura de Ciclo (Días 1 a 5)
1. Carga `Base para el como vamos.xlsx` y `Objetivos Arte.xlsx`.
2. Revisa la meta de **Disponibles Proyectadas** y **Desafíos de Inicios/Reinicios**.
3. Realiza la alineación de metas con cada Líder de Negocio utilizando la pestaña *Cómo Vamos*.

### 🟡 Fase 2: Desarrollo y Mantenimiento (Días 6 a 15)
1. Sincroniza cada 2 días con `mi_grupo.xls` o `activas.xlsx` para registrar los pedidos que van entrando.
2. Abre la pestaña *Gera Crédito & Cobranza*: audita deudas de consultoras con pedido retenido para liberar facturación.
3. Activa la campaña sobre **Inactivas 1 (I1) e Inactivas 2 (I2)** para asegurar recuperos tempranos.

### 🔴 Fase 3: Cierre de Campaña (Últimos 4 días)
1. **Auditoría de Brecha:** Consulta la columna *Falta para el 100%* en la tabla de facturación para direccionar los esfuerzos donde falte poco para alcanzar el escalón de ganancia.
2. **Campaña Salvemos I3:** Contacto telefónico y visita personalizada a las consultoras en riesgo de fuga definitiva (I3).
3. **Asegurar Saldo Comercial Positivo:** Confirmar que Inicios + Reinicios + Recuperos superen las cesantes por al menos $+2$.
4. **Resguardo de Cierre:** En la barra lateral, pestaña Tableau, presionar *"📸 Asegurar Cierre de Ciclo en Histórico"* para congelar los puntos y resultados del ciclo que finaliza.

---
*Documento maestro generado para el ecosistema de gestión comercial `app_lideres`.*
