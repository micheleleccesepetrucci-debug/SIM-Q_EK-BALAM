# Documento de Soporte Técnico: Herramienta de Integridad de Estructuras Civiles (Matriz 3x3)

Este documento técnico detalla la arquitectura, el esquema de datos de entrada y la lógica del motor de cálculo de riesgos del sistema de jerarquización de plataformas costa afuera (Complejo Ek-Balam) ajustado a la **Matriz de Criticidad de 3x3** bajo la norma **API RP 2SIM** (*Structural Integrity Management for Fixed Offshore Structures*).

Su propósito es servir como **contexto de referencia técnica ("Context Window Prompt Support")** para alimentar a modelos de inteligencia artificial (como **Gemini**) y automatizar la redacción de reportes ejecutivos e informes de flota.

---

## 1. Arquitectura General de la Herramienta

La herramienta está estructurada como una aplicación web de cliente único (Single Page Application - SPA) estática, optimizada para procesar datos localmente en el navegador a través de archivos estáticos cargados en el cliente:

```
[Directorio Raíz]
├── app/
│   ├── index.html   --> Interfaz de usuario (KPIs, tablas y contenedores de gráficos SVG).
│   ├── style.css    --> Sistema de diseño visual (Tonal depth, glassmorphism, responsive).
│   ├── app.js       --> Motor de cálculo ("Risk Engine"), cargador de Excel y graficador.
│   └── DESIGN.md    --> Documento de directrices estéticas y del sistema de diseño.
├── Base de datos Plataformas EK Balam.xlsx --> Base de datos precargada con evaluaciones de campo.
├── Datos plataformas Ek-Balam.docx         --> Registro histórico de instalación y metadatos de diseño.
├── preguntas_y_opciones.json               --> Catálogo formalizado en formato JSON de preguntas y opciones.
├── preguntas_y_opciones.xlsx               --> Catálogo formalizado de preguntas en Excel.
└── documento_tecnico_soporte.md            --> Este documento de soporte.
```

---

## 2. Esquema de Datos de Entrada (Input Data Schema)

El motor de cálculo unifica dos fuentes de datos para realizar la jerarquización y el análisis de la flota:

### A. Metadatos Estáticos de Plataforma (Diseño y Capacidad)
Configurados en `app.js` y extraídos del registro histórico de instalación (`Datos plataformas Ek-Balam.docx`) calculando la antigüedad respecto al año corriente **2026**:

*   **Antigüedad (Age):** Años en servicio desde la fecha de instalación (ej. 1988 $\rightarrow$ 38 años en 2026).
*   **Categoría de Exposición Base (Occupancy / CoF):** Clasificación de exposición según la API RP 2SIM (L-1, L-2, L-3).
*   **Volumen de Producción (bpd):** Barriles de petróleo por día manejados.
*   **Personal a bordo (occupancyVal):** Dotación de personal promedio en régimen ordinario.

```javascript
const platformMetadata = {
    'EK-A': { age: 38, occupancy: 'L-2', occupancyVal: 45, bpd: 15000 },
    'EK-A-Hab': { age: 38, occupancy: 'L-1', occupancyVal: 120, bpd: 0 },
    'EK-A Hab': { age: 38, occupancy: 'L-1', occupancyVal: 120, bpd: 0 },
    'EK-A Perf 2': { age: 12, occupancy: 'L-2', occupancyVal: 60, bpd: 5000 },
    'EK-TA': { age: 34, occupancy: 'L-3', occupancyVal: 0, bpd: 4000 },
    'Balam-TA': { age: 5, occupancy: 'L-3', occupancyVal: 0, bpd: 6000 },
    'Balam-TB': { age: 32, occupancy: 'L-3', occupancyVal: 5, bpd: 18000 },
    'Balam-TC': { age: 33, occupancy: 'L-3', occupancyVal: 0, bpd: 0 }, // Estatus especial: Inactiva
    'Balam-TD': { age: 31, occupancy: 'L-3', occupancyVal: 10, bpd: 12000 },
    'Balam-TE': { age: 30, occupancy: 'L-3', occupancyVal: 0, bpd: 2500 },
    'Balam-1': { age: 33, occupancy: 'L-3', occupancyVal: 2, bpd: 1500 }
};
```
> [!NOTE]
> La plataforma **Balam-TC** se encuentra registrada como una **estructura inactiva y fuera de operación** (0 BPD y 0 personal a bordo). En la interfaz se presenta con un indicador visual gris y discontinuo en el gráfico de burbujas, un badge de inactividad en la tabla y un banner de alerta en su vista individual.

### B. Plantilla de Respuestas a Cuestionario (Excel / JSON)
Cada pestaña de la base de datos representa un cuestionario de 16 preguntas agrupadas en módulos normativos:
*   **Módulo A (Categoría de Seguridad Humana / Ocupación):** A1 a A4.
*   **Módulo B (Consecuencias de Falla):** B1 a B6.
*   **Módulo C (Probabilidad de Falla / Condición Estructural):** C1 a C4.
*   **Módulo D (Iniciadores de Evaluación Estructural):** D1 a D2.

---

## 3. Lógica de Homologación de Respuestas

Dado que el checklist de inspección en campo contiene respuestas binarias o ad-hoc, el sistema realiza una traducción automática hacia opciones formalizadas de la API RP 2SIM (opciones `a`, `b`, o `c`):

| ID Pregunta | Módulo | Variable Evaluada | Regla de Homologación (Traducción a Opción API) |
| :---: | :---: | :--- | :--- |
| **A1** | A | Régimen de Ocupación | **a** $\rightarrow$ Habitada continuamente. **b** $\rightarrow$ De trabajo. **c** $\rightarrow$ No habitada (Unmanned). |
| **A2** | A | Plan de Evacuación | **a** $\rightarrow$ Sí (Garantizado). **b** $\rightarrow$ No (Inseguro). **c** $\rightarrow$ No aplica (No habitada). |
| **B1** | B | Fluidos e H2S | **a** $\rightarrow$ Sí y sin ESD. **b** $\rightarrow$ Sí con SSSV operativas. **c** $\rightarrow$ No (Sin flujo autógeno). |
| **B2** | B | Rol de Infraestructura | **a** $\rightarrow$ Complejo Troncal (C-1). **b** $\rightarrow$ Intermedia (C-2). **c** $\rightarrow$ Satélite Mínima (C-3). |
| **C1** | C | Luz Libre / Air Gap | **a** $\rightarrow$ Inadecuada (Wave-in-deck). **b** $\rightarrow$ Marginal ($<1\text{m}$). **c** $\rightarrow$ Adecuada. |
| **C2** | C | Redundancia Jacket | **a** $\rightarrow$ Caisson/Trípode (Baja). **b** $\rightarrow$ 4 patas arriostramiento K. **c** $\rightarrow$ X-bracing (Alta). |
| **C3** | C | Daño Físico / Degradación | **a** $\rightarrow$ Severo ($>10\%$ pérdida). **b** $\rightarrow$ Moderado. **c** $\rightarrow$ Sin daños significativos. |
| **D1** | D | Iniciador Personal | **a** $\rightarrow$ Incremento de POB / cambio categoría. **b** $\rightarrow$ Sin adición de personal. |
| **D2** | D | Iniciador Equipos | **a** $\rightarrow$ Nuevos módulos de proceso/cabezas de pozo. **b** $\rightarrow$ Sin adición de equipos. |

---

## 4. El Motor de Cálculo de Criticidad (Matriz 3x3)

El Risk Engine de la aplicación clasifica cada activo en una matriz bidimensional cruce de **Probabilidad de Falla (LoF)** y **Consecuencia de Falla (CoF)**:

### 1. Categorización de Consecuencias (CoF / Exposure Category)
Determinado a partir de la envolvente de las categorías de ocupación humana (Módulo A) y consecuencias económicas/ambientales (Módulo B):
*   **Categoría de Seguridad Humana (S-1, S-2, S-3):**
    *   Si `A1 === 'c' || A2 === 'c'` $\rightarrow$ **S-3** (Bajo impacto a vida humana).
    *   Si `A2 === 'b'` $\rightarrow$ **S-1** (Manned-nonevacuated, alto impacto).
    *   Si `(A1 === 'a' || A1 === 'b') && A2 === 'a'` $\rightarrow$ **S-2** (Evacuación pre-evento garantizada).
    *   *Excepción Técnica:* Las plataformas habitacionales (`EK-A-Hab` / `EK-A Hab`) se fuerzan a categoría **S-1** debido a la exposición humana permanente.
*   **Categoría de Consecuencia de Falla (C-1, C-2, C-3):**
    *   Si `B1 === 'a' || B2 === 'a'` $\rightarrow$ **C-1** (Alta consecuencia económica/ambiental).
    *   Si `B1 === 'b' || B2 === 'b'` $\rightarrow$ **C-2** (Consecuencia intermedia).
    *   Sino $\rightarrow$ **C-3** (Consecuencia baja).
*   **Exposición Final CoF (L-1, L-2, L-3):**
    *   Si `S-Category === 'S-1'` $\rightarrow$ **L-1** (Alta Consecuencia).
    *   Si `S-Category === 'S-2'`: Si es `C-1` $\rightarrow$ **L-1**; sino $\rightarrow$ **L-2**.
    *   Si `S-Category === 'S-3'`: Si es `C-1` $\rightarrow$ **L-1**; si es `C-2` $\rightarrow$ **L-2**; sino $\rightarrow$ **L-3**.

### 2. Categorización de Probabilidad (LoF)
Evaluada directamente según la condición física estructural del Jacket (Módulo C):
*   **LoF ALTA:** Si `C1 === 'a'` (Inundación de cubierta / Wave-in-deck) **ó** `C3 === 'a'` (Daño estructural severo o fatiga crítica activa).
*   **LoF BAJA:** Si la estructura es óptima en todas las variables críticas: `C1 === 'c'` (Luz libre adecuada) **y** `C2 === 'c'` (Configuración con redundancia de cargas redundantes) **y** `C3 === 'c'` (Sin daño físico).
*   **LoF MEDIA:** Cualquier otra combinación intermedia.

### 3. Cruzamiento y Matriz de Criticidad 3x3
La intersección de la fila de **LoF** y la columna de **CoF (Exposición)** define el nivel de riesgo del activo:

| Probabilidad (LoF) \ Exposición (CoF) | L-1 (Alta Consecuencia) | L-2 (Media Consecuencia) | L-3 (Baja Consecuencia) |
| :---: | :---: | :---: | :---: |
| **ALTA** | **Alto (Nivel 1)** | **Alto (Nivel 1)** | **Medio (Nivel 2)** |
| **MEDIA** | **Alto (Nivel 1)** | **Medio (Nivel 2)** | **Bajo (Nivel 3)** |
| **BAJA** | **Medio (Nivel 2)** | **Bajo (Nivel 3)** | **Bajo (Nivel 3)** |

---

## 5. Iniciadores de Reevaluación Estructural (Gatillos Módulo D)

Se analizan Iniciadores de Evaluación Estructural Detallada (Módulo D) que disparan la necesidad de reevaluación civil estructural:
*   **D1:** Incremento de personal (POB).
*   **D2:** Adición de equipos/módulos.
*   **D3:** Incremento de cargas por sobrepeso.
*   **D4:** Nivel extremo de ola tocando cubierta.
*   **D5:** Daño físico acumulado en elementos críticos.

Si **cualquiera** de estas variables se encuentra activa (respuesta `a`), se enciende el banner de **⚠️ GATILLO DE EVALUACIÓN ACTIVO (ANÁLISIS DETALLADO REQUERIDO)** indicando la necesidad de escalar la plataforma a análisis estructural detallado.

---

## 6. Salidas de Inspección Recomendadas

Según la clasificación de riesgo obtenida, se establecen los intervalos límite máximos permitidos para realizar las campañas de inspección submarina y aérea (NRF-260-PEMEX / API RP 2SIM):

| Riesgo Resultante | GVI (Inspección Visual General) | CVI (Visual Detallada / Ánodos) | NDT (Ensayos No Destructivos / ACFM) |
| :---: | :---: | :---: | :--- |
| **Alto (Nivel 1)** | Cada 1 - 2 años | Cada 3 - 5 años (Dirigida) | Cada 3 - 5 años (ACFM/FMD en nodos críticos) |
| **Medio (Nivel 2)** | Cada 3 - 5 años | Cada 6 - 10 años | Subacuático bajo monitoreo dirigido |
| **Bajo (Nivel 3)** | Cada 5 - 6 años | Cada 10 - 12 años | Según campaña regular del activo |

---

## 7. Módulos Gráficos y Analítica de Flota

La herramienta procesa y visualiza los indicadores a través de tres vistas en la pestaña **Jerarquización General**:
1.  **Cuadrante de Decisión (Burbujas):** Posiciona las plataformas cruzando los años en servicio (Eje X, 0 a 40 años) y su LoF cualitativo (Eje Y). El tamaño de la burbuja refleja su volumen de POB y el color su nivel de riesgo.
2.  **Pareto de Desviaciones:** Jerarquiza los criterios con mayor recurrencia de fallas (opciones no conformes), lo que permite planificar programas de mantenimiento enfocados en las causas de riesgo raíz comunes de la flota.
3.  **Flujo de Inspección (Sankey):** Mapea de forma dinámica la distribución de plataformas desde su estado actual de criticidad hasta los intervalos recomendados de GVI, CVI y NDT.
