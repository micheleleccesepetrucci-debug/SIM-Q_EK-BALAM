// Base de datos de plataformas precargada desde el Excel "Base de datos Plataformas EK Balam.xlsx"
// Respuestas mapeadas directamente a los módulos normativos API RP 2SIM / DNV-RP-C210 / ISO 19902
const defaultExcelData = {
  "EK-A": {
    "A1": "b",
    "A2": "a",
    "B1": "b",
    "B2": "a",
    "C1": "b",
    "C2": "b",
    "C3": "c",
    "D1": "a",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "EK-A Perf 2": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "b",
    "C1": "c",
    "C2": "c",
    "C3": "c",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "EK-TA": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "b",
    "C1": "c",
    "C2": "c",
    "C3": "c",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "Balam-TA": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "b",
    "C1": "c",
    "C2": "b",
    "C3": "c",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "Balam-TB": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "b",
    "C1": "c",
    "C2": "b",
    "C3": "b",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "Balam-TD": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "b",
    "C1": "c",
    "C2": "b",
    "C3": "c",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "Balam-TE": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "b",
    "C1": "c",
    "C2": "c",
    "C3": "c",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "Balam-1": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "c",
    "C1": "c",
    "C2": "b",
    "C3": "c",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "Balam-TC": {
    "A1": "c",
    "A2": "c",
    "B1": "c",
    "B2": "c",
    "C1": "c",
    "C2": "a",
    "C3": "b",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "EK-A Hab": {
    "A1": "a",
    "A2": "a",
    "B1": "c",
    "B2": "c",
    "C1": "c",
    "C2": "c",
    "C3": "c",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "b"
  },
  "Balam-A": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "b",
    "C1": "c",
    "C2": "c",
    "C3": "c",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "b",
    "D5": "b"
  },
  "EK-TB": {
    "A1": "c",
    "A2": "c",
    "B1": "b",
    "B2": "c",
    "C1": "a",
    "C2": "a",
    "C3": "a",
    "D1": "b",
    "D2": "b",
    "D3": "b",
    "D4": "a",
    "D5": "a"
  }
};

// Estructura de Metadatos de Preguntas alineadas a API RP 2SIM
const questionsMetadata = [
    // MÓDULO A: CATEGORÍA DE SEGURIDAD HUMANA / OCUPACIÓN
    {
        id: 'A1',
        module: 'A',
        moduleName: { es: 'Módulo A: Categoría de Seguridad Humana / Ocupación', en: 'Module A: Life Safety Category' },
        type: 'consequence',
        isPrincipal: false,
        text: {
            es: 'Régimen de Ocupación: ¿Cuál es el régimen habitual de ocupación de personal en la plataforma?',
            en: 'Occupancy Regime: What is the normal occupancy regime of personnel on the platform?'
        },
        label: {
            es: 'A.1 - Régimen de Ocupación',
            en: 'A.1 - Occupancy Regime'
        },
        options: {
            a: {
                es: 'Habitada continuamente / Habitacional permanente',
                en: 'Continuously manned / Permanent living quarters'
            },
            b: {
                es: 'Habitada operativamente / De trabajo (evacuación planificada)',
                en: 'Operationally manned / Working (planned evacuation)'
            },
            c: {
                es: 'No habitada (Unmanned) / Ocasional',
                en: 'Unmanned / Occasional visits only'
            }
        },
        mitigationAction: {
            es: 'Automatizar las operaciones para operar de forma no tripulada (unmanned) y reducir la dotación.',
            en: 'Automate operations to run unmanned and reduce personnel on board.'
        }
    },
    {
        id: 'A2',
        module: 'A',
        moduleName: { es: 'Módulo A: Categoría de Seguridad Humana / Ocupación', en: 'Module A: Life Safety Category' },
        type: 'consequence',
        isPrincipal: false,
        text: {
            es: 'Plan de Evacuación: ¿Se cuenta con un plan de evacuación probado, tiempo suficiente y recursos aéreos/marítimos en huracán/Norte?',
            en: 'Evacuation Plan: Is there a proven evacuation plan, sufficient warning time, and air/sea resources for hurricanes?'
        },
        label: {
            es: 'A.2 - Plan de Evacuación',
            en: 'A.2 - Evacuation Plan'
        },
        options: {
            a: {
                es: 'Sí: Plan formal y recursos logísticos plenamente garantizados (Clasifica como S-2)',
                en: 'Yes: Formal plan and fully guaranteed logistical resources (Classifies as S-2)'
            },
            b: {
                es: 'No: Evacuación pre-evento no garantizada o impracticable (Clasifica como S-1)',
                en: 'No: Pre-event evacuation not guaranteed or impracticable (Classifies as S-1)'
            },
            c: {
                es: 'No aplica: Plataforma no habitada (Clasifica como S-3)',
                en: 'Not applicable: Unmanned platform (Classifies as S-3)'
            }
        },
        mitigationAction: {
            es: 'Desarrollar y certificar un protocolo formal de evacuación pre-tormenta con logística de PEMEX.',
            en: 'Develop and certify a formal pre-storm evacuation protocol with PEMEX logistics.'
        }
    },
    // MÓDULO B: CONSECUENCIAS DE FALLA
    {
        id: 'B1',
        module: 'B',
        moduleName: { es: 'Módulo B: Consecuencias de Falla', en: 'Module B: Consequence of Failure Category' },
        type: 'consequence',
        isPrincipal: false,
        text: {
            es: 'Riesgo por Fluido y Gas Amargo H2S: ¿Maneja o está conectada a pozos con flujo autógeno de H2S o aceite?',
            en: 'Fluid and Sour Gas H2S Risk: Does it handle or connect to wells with natural flowing H2S or oil?'
        },
        label: {
            es: 'B.1 - Fluidos e H2S',
            en: 'B.1 - Fluids and H2S'
        },
        options: {
            a: {
                es: 'Sí, y el cierre automático no está garantizado o es un nodo troncal mayor (Clasifica como C-1)',
                en: 'Yes, and automatic shutdown is not guaranteed or it is a major trunk node (Classifies as C-1)'
            },
            b: {
                es: 'Sí, pero cuenta con Válvulas de Seguridad Subsuperficiales (SSSV) operativas e inspeccionadas (Clasifica como C-2)',
                en: 'Yes, but has operational and inspected Subsurface Safety Valves (SSSV) (Classifies as C-2)'
            },
            c: {
                es: 'No: No maneja hidrocarburos o pozos con bombeo mecánico/neumático sin flujo autógeno (Clasifica como C-3)',
                en: 'No: Does not handle hydrocarbons or wells with mechanical/pneumatic lift, no natural flow (Classifies as C-3)'
            }
        },
        mitigationAction: {
            es: 'Instalar o programar pruebas preventivas de las válvulas SSSV e instrumentación de aislamiento rápido.',
            en: 'Install or schedule preventive testing of SSSV valves and quick isolation instrumentation.'
        }
    },
    {
        id: 'B2',
        module: 'B',
        moduleName: { es: 'Módulo B: Consecuencias de Falla', en: 'Module B: Consequence of Failure Category' },
        type: 'consequence',
        isPrincipal: false,
        text: {
            es: 'Rol de Infraestructura: ¿Cuál es el rol operativo y de infraestructura en la red de la Sonda de Campeche?',
            en: 'Infrastructure Role: What is the operational and infrastructure role in the Campeche Sound network?'
        },
        label: {
            es: 'B.2 - Rol de Infraestructura',
            en: 'B.2 - Infrastructure Role'
        },
        options: {
            a: {
                es: 'Infraestructura Crítica / Complejo Troncal: Exportación, compresión mayor o almacenamiento (Clasifica como C-1)',
                en: 'Critical Infrastructure / Major Trunk Node: Export, major compression, or storage (Classifies as C-1)'
            },
            b: {
                es: 'Plataforma Intermedia / Producción Estándar: Proceso/perforación aislable pre-evento (Clasifica como C-2)',
                en: 'Intermediate Platform / Standard Production: Process/drilling isolatable pre-event (Classifies as C-2)'
            },
            c: {
                es: 'Plataforma Mínima / Satélite: Estructura pequeña, ≤5 pozos, tirante ≤30m, cubierta ≤37m² (Clasifica como C-3)',
                en: 'Minimal Platform / Satellite: Small structure, ≤5 wells, water depth ≤30m, deck area ≤37m² (Classifies as C-3)'
            }
        },
        mitigationAction: {
            es: 'Implementar loops de desvío o interconexiones que permitan aislar la plataforma sin paralizar la flota.',
            en: 'Implement bypass loops or interconnects to isolate the platform without disabling the fleet.'
        }
    },
    // MÓDULO C: PROBABILIDAD DE FALLA / CONDICIÓN ESTRUCTURAL
    {
        id: 'C1',
        module: 'C',
        moduleName: { es: 'Módulo C: Probabilidad de Falla / Condición Estructural', en: 'Module C: Likelihood of Failure' },
        type: 'probability',
        isPrincipal: true,
        text: {
            es: 'Luz Libre / Air Gap: ¿Cuál es la elevación de la cubierta Cellar Deck respecto a la cresta de ola metoceánica extrema?',
            en: 'Air Gap: What is the elevation of the Cellar Deck relative to the extreme metocean wave crest?'
        },
        label: {
            es: 'C.1 - Luz Libre / Air Gap',
            en: 'C.1 - Air Gap status'
        },
        options: {
            a: {
                es: 'Inadecuada / Impacto de ola en cubierta (Wave-in-deck) (Dispara Probabilidad ALTA)',
                en: 'Inadequate / Wave-in-deck impact (Triggers HIGH Likelihood)'
            },
            b: {
                es: 'Luz libre (Air Gap) Marginal: Margen menor a 1 metro',
                en: 'Marginal Air Gap: Margin of clearance less than 1 meter'
            },
            c: {
                es: 'Luz libre (Air Gap) Adecuada: Cumple con la elevación de seguridad sin riesgo de inundación',
                en: 'Adequate Air Gap: Meets safety clearance elevation without wave impact risk'
            }
        },
        mitigationAction: {
            es: 'Aligerar cubiertas inferiores, remover obstrucciones o realizar análisis hidrodinámico de ola en cubierta.',
            en: 'Clear lower decks, remove obstructions, or perform wave-in-deck hydrodynamic analysis.'
        }
    },
    {
        id: 'C2',
        module: 'C',
        moduleName: { es: 'Módulo C: Probabilidad de Falla / Condición Estructural', en: 'Module C: Likelihood of Failure' },
        type: 'probability',
        isPrincipal: false,
        text: {
            es: 'Redundancia Jacket: ¿Qué tipo de configuración geométrica de arriostramiento (Bracing) posee la subestructura?',
            en: 'Jacket Redundancy: What type of bracing geometric configuration does the substructure have?'
        },
        label: {
            es: 'C.2 - Redundancia Jacket',
            en: 'C.2 - Jacket Redundancy'
        },
        options: {
            a: {
                es: 'Baja Redundancia / Crítica: Caisson, Trípode, o 4 piernas con arriostramiento K/diagonal sin caminos alternos',
                en: 'Low Redundancy: Caisson, Tripod, or 4 legs with K/diagonal bracing and no alternate load paths'
            },
            b: {
                es: 'Redundancia Moderada: Estructura de 4 o 6 piernas con arriostramiento diagonal combinado',
                en: 'Moderate Redundancy: 4 or 6 legs structure with combined diagonal bracing'
            },
            c: {
                es: 'Alta Redundancia / Robusta: Estructura de 8 piernas o más con arriostramiento tipo X (X-braced)',
                en: 'High Redundancy: 8 legs or more structure with X-bracing (multiple load paths)'
            }
        },
        mitigationAction: {
            es: 'Realizar estudios estructurales de capacidad remanente e instalar abrazaderas de refuerzo en nudos críticos.',
            en: 'Perform structural residual strength studies and install reinforcing clamps at critical joints.'
        }
    },
    {
        id: 'C3',
        module: 'C',
        moduleName: { es: 'Módulo C: Probabilidad de Falla / Condición Estructural', en: 'Module C: Likelihood of Failure' },
        type: 'probability',
        isPrincipal: true,
        text: {
            es: 'Daño Físico o Degradación: ¿Cuál es el estado de daño mecánico, grietas por fatiga o corrosión de la estructura?',
            en: 'Physical Damage or Degradation: What is the status of mechanical damage, fatigue cracks, or corrosion?'
        },
        label: {
            es: 'C.3 - Daño Físico / Degradación',
            en: 'C.3 - Structural Damage / Degradation'
        },
        options: {
            a: {
                es: 'Daño Severo / Reducción >10% capacidad: Miembros rotos/faltantes, grietas críticas o corrosión severa (ALTA)',
                en: 'Severe Damage / >10% capacity reduction: Broken/missing members, critical cracks, or severe corrosion (HIGH)'
            },
            b: {
                es: 'Daño Moderado: Corrosión splash zone localizada, abolladuras menores o pérdida parcial de ánodos',
                en: 'Moderate Damage: Localized splash zone corrosion, minor dents, or partial depletion of anodes'
            },
            c: {
                es: 'Sin Daños Significativos: Estructura en buen estado general, protección catódica conforme (≤-800mV)',
                en: 'No Significant Damage: Substructure in good general condition, compliant cathodic protection (≤-800mV)'
            }
        },
        mitigationAction: {
            es: 'Ejecutar reparaciones submarinas (soldadura, amolado de grietas o relleno con lechada) y retrofit de ánodos.',
            en: 'Execute subsea repairs (welding, crack grinding, or grout filling) and anode retrofit.'
        }
    },
    // MÓDULO D: INICIADORES DE EVALUACIÓN (ASSESSMENT INITIATORS)
    {
        id: 'D1',
        module: 'D',
        moduleName: { es: 'Módulo D: Iniciadores de Evaluación Estructural Detallada', en: 'Module D: Assessment Initiators' },
        type: 'probability',
        isPrincipal: false,
        text: {
            es: 'Iniciador 1 (Personal): ¿Se ha incrementado la dotación de personal o cambiado a una categoría más restrictiva?',
            en: 'Initiator 1 (Personnel): Has personnel POB increased or changed to a more restrictive life safety category?'
        },
        label: {
            es: 'D.1 - Iniciador Personal',
            en: 'D.1 - Personnel Initiator'
        },
        options: {
            a: {
                es: 'Sí: Dispara Reevaluación Estructural',
                en: 'Yes: Triggers Structural Reevaluation'
            },
            b: {
                es: 'No: Sin incrementos ni cambios de categoría',
                en: 'No: No POB increase or category changes'
            }
        },
        mitigationAction: {
            es: 'Reducir la dotación para volver a la categoría de diseño original o realizar estudio de extensión de vida.',
            en: 'Reduce POB to revert to original design category or perform life extension study.'
        }
    },
    {
        id: 'D2',
        module: 'D',
        moduleName: { es: 'Módulo D: Iniciadores de Evaluación Estructural Detallada', en: 'Module D: Assessment Initiators' },
        type: 'probability',
        isPrincipal: false,
        text: {
            es: 'Iniciador 2 (Instalaciones): ¿Se han instalado nuevos módulos, pozos o equipos que aumenten las consecuencias?',
            en: 'Initiator 2 (Facilities): Have new modules, wells, or equipment been added that increase consequence category?'
        },
        label: {
            es: 'D.2 - Iniciador Instalaciones',
            en: 'D.2 - Facilities Initiator'
        },
        options: {
            a: {
                es: 'Sí: Dispara Reevaluación Estructural',
                en: 'Yes: Triggers Structural Reevaluation'
            },
            b: {
                es: 'No: Sin adiciones de equipos mayores',
                en: 'No: No major equipment additions'
            }
        },
        mitigationAction: {
            es: 'Realizar análisis de flexibilidad e interconexión para justificar la exención de reevaluación o retirar equipos inactivos.',
            en: 'Perform flexibility analysis to justify reevaluation waiver or decommission inactive equipment.'
        }
    },
    {
        id: 'D3',
        module: 'D',
        moduleName: { es: 'Módulo D: Iniciadores de Evaluación Estructural Detallada', en: 'Module D: Assessment Initiators' },
        type: 'probability',
        isPrincipal: false,
        text: {
            es: 'Iniciador 3 (Sobrecarga): ¿Existe un incremento de cargas operativas o ambientales >10% respecto al diseño?',
            en: 'Initiator 3 (Overload): Is there an increase in operating or environmental loads >10% compared to design?'
        },
        label: {
            es: 'D.3 - Iniciador Sobrecarga',
            en: 'D.3 - Overload Initiator'
        },
        options: {
            a: {
                es: 'Sí: Dispara Reevaluación Estructural',
                en: 'Yes: Triggers Structural Reevaluation'
            },
            b: {
                es: 'No: Cargas operativas y ambientales controladas dentro de límites',
                en: 'No: Operating and environmental loads controlled within limits'
            }
        },
        mitigationAction: {
            es: 'Desarrollar una campaña de reducción de peso (aligerar topside) e instrumentar la estructura.',
            en: 'Execute a weight reduction campaign (lighten topside) and instrument the structure.'
        }
    },
    {
        id: 'D4',
        module: 'D',
        moduleName: { es: 'Módulo D: Iniciadores de Evaluación Estructural Detallada', en: 'Module D: Assessment Initiators' },
        type: 'probability',
        isPrincipal: false,
        text: {
            es: 'Iniciador 4 (Ola en Cubierta): ¿La cresta de ola extrema alcanza o supera la estructura inferior (Cellar Deck)?',
            en: 'Initiator 4 (Wave-in-deck): Does the extreme wave crest reach or exceed the lower deck (Cellar Deck)?'
        },
        label: {
            es: 'D.4 - Iniciador Ola en Cubierta',
            en: 'D.4 - Wave-in-deck Initiator'
        },
        options: {
            a: {
                es: 'Sí: Dispara Reevaluación Estructural de Resistencia Última (Pushover)',
                en: 'Yes: Triggers Ultimate Strength Structural Reevaluation (Pushover)'
            },
            b: {
                es: 'No: Margen de luz libre conforme',
                en: 'No: Air gap margin compliant'
            }
        },
        mitigationAction: {
            es: 'Realizar un análisis Pushover Nivel III con fuerzas hidrodinámicas de ola en cubierta y mitigar componentes.',
            en: 'Perform a Level III Pushover analysis with hydrodynamic wave-in-deck loads and mitigate components.'
        }
    },
    {
        id: 'D5',
        module: 'D',
        moduleName: { es: 'Módulo D: Iniciadores de Evaluación Estructural Detallada', en: 'Module D: Assessment Initiators' },
        type: 'probability',
        isPrincipal: false,
        text: {
            es: 'Iniciador 5 (Daño Acumulado): ¿Existen daños estructurales o deterioro que reduzcan la capacidad global en ≥10%?',
            en: 'Initiator 5 (Accumulated Damage): Is there structural damage or degradation reducing global capacity by ≥10%?'
        },
        label: {
            es: 'D.5 - Iniciador Daño Acumulado',
            en: 'D.5 - Damage Initiator'
        },
        options: {
            a: {
                es: 'Sí: Dispara Reevaluación Estructural obligatoria',
                en: 'Yes: Triggers mandatory Structural Reevaluation'
            },
            b: {
                es: 'No: Deterioro acumulado por debajo del límite de tolerancia global',
                en: 'No: Accumulated degradation below global tolerance limit'
            }
        },
        mitigationAction: {
            es: 'Reparar los elementos deformados o agrietados inmediatamente para restablecer la capacidad de diseño.',
            en: 'Repair buckled or cracked members immediately to restore design capacity.'
        }
    }
];

// Definición de Niveles de Riesgo (3 Niveles API RP 2SIM)
const riskLevels = {
    'L': { name: { es: 'Bajo (Nivel 3)', en: 'Low (Level 3)' }, class: 'risk-low', color: '#05A405' },
    'M': { name: { es: 'Medio (Nivel 2)', en: 'Medium (Level 2)' }, class: 'risk-medium', color: '#FFDE13' },
    'H': { name: { es: 'Alto (Nivel 1)', en: 'High (Level 1)' }, class: 'risk-high', color: '#D02309' }
};

// Matriz de Criticidad (3x3) basada en Categoría de Exposición CoF (Columnas: L-1, L-2, L-3)
// y Categoría de Probabilidad de Falla LoF (Filas: ALTA, MEDIA, BAJA)
const matrixValues = {
    'ALTA':  { 'L-1': 'H', 'L-2': 'H', 'L-3': 'M' }, // Row ALTA: L-1 (Alto), L-2 (Alto), L-3 (Medio)
    'MEDIA': { 'L-1': 'H', 'L-2': 'M', 'L-3': 'L' }, // Row MEDIA: L-1 (Alto), L-2 (Medio), L-3 (Bajo)
    'BAJA':  { 'L-1': 'M', 'L-2': 'L', 'L-3': 'L' }  // Row BAJA: L-1 (Medio), L-2 (Bajo), L-3 (Bajo)
};

// Diccionario de Traducción de la Interfaz de Usuario
const uiTranslations = {
    es: {
        title: "SIM-Q RELIARISK",
        subtitle: "Sistema de Gestión de Integridad Estructural - Complejo Ek-Balam",
        dragText: 'Arrastra aquí "Base de datos Plataformas EK Balam.xlsx"',
        clickText: "o haz clic para buscar el archivo",
        noFile: "No se ha seleccionado ningún archivo",
        loadedFile: "Archivo cargado",
        tabDashboard: "📊 Jerarquización General",
        kpiTotal: "Total Plataformas",
        kpiVeryHigh: "", // deprecated in 3x3
        kpiHigh: "Alto Riesgo (N1)",
        kpiMedium: "Medio Riesgo (N2)",
        kpiLow: "Bajo Riesgo (N3)",
        matrixTitle: "Matriz de Criticidad Consolidada (3x3)",
        matrixSubtitle: "Posición relativa de todas las instalaciones evaluadas (API RP 2SIM)",
        tableTitle: "Jerarquización de Plataformas por Criticidad",
        tableSubtitle: "Listado priorizado de mayor a menor riesgo estructural",
        colPlat: "Plataforma",
        colBaseCons: "Exposición (CoF)",
        colPoints: "Probabilidad (LoF)",
        colFinalCons: "Exposición CoF",
        colCoord: "Coordenada",
        colRisk: "Riesgo",
        colAction: "Acción",
        btnInspect: "Simular",
        platDescDefault: "Plataforma Satélite de Extracción",
        platDescEKA: "Nodo Central de Control, Generación Eléctrica y Concentración",
        platDescEKAHab: "Módulo Habitacional y Telecomunicaciones del Activo",
        platDescEKAPerf: "Plataforma de Inyección de Agua y Soporte de Presión",
        platDescBalamTB: 'Cabecera del Colector Sur (Ducto de 20")',
        platDescBalamTD: "Cabecera del Colector Norte y Planta de Inyección",
        platTitleQuestions: "Cuestionario de Parámetros API RP 2SIM (Nivel I)",
        platSubtitleQuestions: "Modifica las respuestas para simular el análisis de riesgos",
        lblBaseCons: "Categoría de Ocupación Humana (Módulo A)",
        platTitleMatrix: "Matriz de Criticidad de la Plataforma (3x3)",
        platTitleStats: "Resultados de Simulación",
        platTitleMitigation: "Plan de Mitigación Recomendado",
        platSubtitleMitigation: "Acciones requeridas para llevar la plataforma a un nivel de <strong>Riesgo Bajo (Nivel 3)</strong>",
        statProbPoints: "Score Condición",
        statProbLevel: "Probabilidad LoF",
        statConsFinal: "Exposición CoF",
        originalLabel: "Original",
        successOptimal: "<strong>🏆 Plataforma en estado óptimo:</strong> Esta instalación no presenta ningún iniciador activo ni deficiencia estructural. Su nivel de riesgo es Bajo (Nivel 3).",
        successGoal: "<strong>✔️ Objetivo Logrado (Riesgo Bajo):</strong> Al simular la aplicación de estas medidas, la plataforma alcanza el nivel de <strong>Riesgo Bajo (Nivel 3)</strong>.",
        reqTextCons3: "Se requiere mitigar <strong>todas</strong> las deficiencias de probabilidad activas.",
        reqTextCons2: "Se requiere mitigar la probabilidad a <strong>Baja</strong>.",
        reqTextCons1: "Se requiere mitigar la probabilidad a <strong>Media</strong> o <strong>Baja</strong>.",
        bannerWarning: "<strong>⚠️ Nivel de Riesgo Simulado: {riskName} ({coord}).</strong><br>Para clasificar la plataforma como <strong>Riesgo Bajo (Nivel 3)</strong>, {reqText}",
        deficiencyLbl: "Deficiencia",
        actionLbl: "Acción Recomendada",
        triggerBadge: "⚠️ GATILLO DE EVALUACIÓN ACTIVO (ANÁLISIS DETALLADO REQUERIDO)",
        inspectionTitle: "Intervalos de Inspección Recomendados (API RP 2SIM)",
        gviLabel: "GVI (Visual General)",
        cviLabel: "CVI (Visual Detallada)",
        ndtLabel: "NDT (Ensayos No Dest.)",
        lblAnalyticsTitle: "Análisis e Indicadores de Flota",
        lblAnalyticsSubtitle: "Visualizaciones estratégicas del Complejo Ek-Balam",
        tabBubble: "🫧 Cuadrante de Decisión (Burbujas)",
        tabPareto: "📊 Pareto de Desviaciones",
        tabSankey: "🔀 Flujo de Inspección (Sankey)",
        platformChartsTitle: "Análisis Gráfico de Integridad",
        platformChartsSubtitle: "Causas de riesgo y perfil de salud estructural (API RP 2SIM)",
        lblRadarChart: "Perfil de Salud (Radar)",
        lblWaterfallChart: "Descomposición de Riesgo (Cascada)",
        tooltipAge: "Antigüedad",
        tooltipLoF: "Nivel LoF",
        tooltipCoF: "Nivel CoF",
        tooltipRisk: "Nivel Riesgo",
        tooltipBpd: "Producción (BPD)",
        tooltipStaff: "Personal a Bordo",
        tooltipCount: "Incidencia",
        tooltipPercentage: "Porcentaje Acumulado",
        btnExportFleet: "📄 Descargar Reporte Word",
        btnExportPlatform: "📄 Descargar Reporte Simulación",
        
        // Suite Layout additions
        navSimulatorName: "🎛️ Simulador por Plataforma",
        navAdminName: "👥 Gestión Usuarios",
        lblSidebarSub: "Sistema de Integridad de Estructuras Civiles",
        lblActiveUser: "USUARIO ACTIVO",
        lblAnalysisMgmt: "Gestión de Análisis",
        lblLoadJsonBtn: "Cargar Archivo (.json)",
        lblExcelOps: "Operaciones de Excel",
        uploadZoneTitle: "Cargar Base de Datos (.xlsx)",
        uploadZoneSubtitle: "Arrastra o haz clic aquí",
        lblFooterSub: "SIM-Q RELIARISK v1.0",
        lblSelectPlatform: "Seleccionar Plataforma a Simular:",
        btnSaveAnalysis: "💾 Guardar Análisis (.json)",
        btnDownloadAnalysis: "⬇️ Descargar (.json)"
    },
    en: {
        title: "SIM-Q RELIARISK",
        subtitle: "Structural Integrity Management System - Ek-Balam Complex",
        dragText: 'Drag "Base de datos Plataformas EK Balam.xlsx" here',
        clickText: "or click to browse file",
        noFile: "No file selected",
        loadedFile: "Loaded file",
        tabDashboard: "📊 General Hierarchy",
        kpiTotal: "Total Platforms",
        kpiVeryHigh: "", // deprecated in 3x3
        kpiHigh: "High Risk (N1)",
        kpiMedium: "Medium Risk (N2)",
        kpiLow: "Low Risk (N3)",
        matrixTitle: "Consolidated Criticality Matrix (3x3)",
        matrixSubtitle: "Relative position of all assessed installations (API RP 2SIM)",
        tableTitle: "Platforms Ranked by Criticality",
        tableSubtitle: "Prioritized list from highest to lowest structural risk",
        colPlat: "Platform",
        colBaseCons: "Exposure (CoF)",
        colPoints: "Likelihood (LoF)",
        colFinalCons: "CoF Exposure",
        colCoord: "Coordinate",
        colRisk: "Risk",
        colAction: "Action",
        btnInspect: "Simulate",
        platDescDefault: "Satellite Extraction Platform",
        platDescEKA: "Central Control, Power Generation and Concentration Node",
        platDescEKAHab: "Living Quarters and Asset Telecommunications Module",
        platDescEKAPerf: "Water Injection and Pressure Support Platform",
        platDescBalamTB: 'Header of South Collector (20" Pipeline)',
        platDescBalamTD: "Header of North Collector and Water Injection Plant",
        platTitleQuestions: "API RP 2SIM Parameters Questionnaire (Level I)",
        platSubtitleQuestions: "Modify answers to simulate a risk analysis",
        lblBaseCons: "Human Occupancy Category (Module A)",
        platTitleMatrix: "Platform Criticality Matrix (3x3)",
        platTitleStats: "Simulation Results",
        platTitleMitigation: "Recommended Mitigation Plan",
        platSubtitleMitigation: "Actions required to bring the platform to a <strong>Low Risk (Level 3)</strong> level",
        statProbPoints: "Condition Score",
        statProbLevel: "LoF Likelihood",
        statConsFinal: "CoF Exposure",
        originalLabel: "Original",
        successOptimal: "<strong>🏆 Platform in optimal state:</strong> This installation does not present any active initiator or structural deficiency. Its risk level is Low (Level 3).",
        successGoal: "<strong>✔️ Target Achieved (Low Risk):</strong> By simulating these actions, the platform reaches a <strong>Low Risk (Level 3)</strong> level.",
        reqTextCons3: "It is required to mitigate <strong>all</strong> active likelihood deficiencies.",
        reqTextCons2: "It is required to reduce likelihood to <strong>Low</strong>.",
        reqTextCons1: "It is required to reduce likelihood to <strong>Medium</strong> or <strong>Low</strong>.",
        bannerWarning: "<strong>⚠️ Simulated Risk Level: {riskName} ({coord}).</strong><br>To classify the platform as <strong>Low Risk (Level 3)</strong>, {reqText}",
        deficiencyLbl: "Deficiency",
        actionLbl: "Recommended Action",
        triggerBadge: "⚠️ ACTIVE ASSESSMENT TRIGGER (DETAILED ANALYSIS REQUIRED)",
        inspectionTitle: "Recommended Inspection Intervals (API RP 2SIM)",
        gviLabel: "GVI (General Visual)",
        cviLabel: "CVI (Close Visual)",
        ndtLabel: "NDT (Non-Destructive)",
        lblAnalyticsTitle: "Fleet Analysis & Indicators",
        lblAnalyticsSubtitle: "Strategic visualizations of the Ek-Balam Complex",
        tabBubble: "🫧 Decision Quadrant (Bubbles)",
        tabPareto: "📊 Pareto of Deviations",
        tabSankey: "🔀 Inspection Flow (Sankey)",
        platformChartsTitle: "Structural Integrity Graphical Analysis",
        platformChartsSubtitle: "Risk drivers and structural health profile (API RP 2SIM)",
        lblRadarChart: "Health Profile (Radar)",
        lblWaterfallChart: "Risk Decomposition (Waterfall)",
        tooltipAge: "Age",
        tooltipLoF: "LoF Level",
        tooltipCoF: "CoF Level",
        tooltipRisk: "Risk Level",
        tooltipBpd: "Production (BPD)",
        tooltipStaff: "Personnel on Board",
        tooltipCount: "Incidence",
        tooltipPercentage: "Cumulative Percentage",
        btnExportFleet: "📄 Download Word Report",
        btnExportPlatform: "📄 Download Simulation Report",
        
        // Suite Layout additions
        navSimulatorName: "🎛️ Platform Simulator",
        navAdminName: "👥 User Management",
        lblSidebarSub: "Civil Structural Integrity System",
        lblActiveUser: "ACTIVE USER",
        lblAnalysisMgmt: "Analysis Management",
        lblLoadJsonBtn: "Load File (.json)",
        lblExcelOps: "Excel Operations",
        uploadZoneTitle: "Load Database (.xlsx)",
        uploadZoneSubtitle: "Drag or click here",
        lblFooterSub: "SIM-Q RELIARISK v1.0",
        lblSelectPlatform: "Select Platform to Simulate:",
        btnSaveAnalysis: "💾 Save Analysis (.json)",
        btnDownloadAnalysis: "⬇️ Download (.json)"
    }
};// Descripciones personalizadas por plataforma
const platformDescriptions = {
    'EK-A': { es: uiTranslations.es.platDescEKA, en: uiTranslations.en.platDescEKA },
    'EK-A-Hab': { es: uiTranslations.es.platDescEKAHab, en: uiTranslations.en.platDescEKAHab },
    'EK-A Hab': { es: uiTranslations.es.platDescEKAHab, en: uiTranslations.en.platDescEKAHab },
    'Balam-TB': { es: uiTranslations.es.platDescBalamTB, en: uiTranslations.en.platDescBalamTB },
    'Balam-TD': { es: uiTranslations.es.platDescBalamTD, en: uiTranslations.en.platDescBalamTD },
    'EK-A Perf 2': { es: uiTranslations.es.platDescEKAPerf, en: uiTranslations.en.platDescEKAPerf },
    'Balam-TC': {
        es: "Estructura inactiva y fuera de operación (producción a 0 BPD y sin personal).",
        en: "Inactive structure and out of operation (0 BPD production and unmanned)."
    },
    'Balam-A': {
        es: "Octópodo de Perforación y Recolección. Aloja la Planta Principal de Tratamiento e Inyección de Agua de Mar.",
        en: "Drilling and Gathering 8-Legged Platform. Houses the Main Seawater Treatment and Injection Plant."
    },
    'EK-TB': {
        es: "Trípode adosado de perforación y producción. Aloja el circuito cerrado de agua congénita.",
        en: "Attached drilling and production Tripod Platform. Houses the closed-loop produced water circuit."
    }
};

// Metadatos de diseño e impacto de las plataformas para análisis y gráficos complementarios
const platformMetadata = {
    'EK-A': { age: 38, occupancy: 'L-2', occupancyVal: 45, bpd: 15000 },
    'EK-A-Hab': { age: 17, occupancy: 'L-1', occupancyVal: 120, bpd: 0 },
    'EK-A Hab': { age: 17, occupancy: 'L-1', occupancyVal: 120, bpd: 0 },
    'EK-A Perf 2': { age: 12, occupancy: 'L-2', occupancyVal: 60, bpd: 5000 },
    'EK-TA': { age: 33, occupancy: 'L-3', occupancyVal: 0, bpd: 4000 },
    'Balam-TA': { age: 33, occupancy: 'L-3', occupancyVal: 0, bpd: 6000 },
    'Balam-TB': { age: 33, occupancy: 'L-3', occupancyVal: 5, bpd: 18000 },
    'Balam-TC': { age: 33, occupancy: 'L-3', occupancyVal: 0, bpd: 0 }, // Inactiva, 0 BPD, 0 Personal
    'Balam-TD': { age: 33, occupancy: 'L-3', occupancyVal: 10, bpd: 12000 },
    'Balam-TE': { age: 33, occupancy: 'L-3', occupancyVal: 0, bpd: 2500 },
    'Balam-1': { age: 34, occupancy: 'L-3', occupancyVal: 2, bpd: 1500 },
    'Balam-A': { age: 12, occupancy: 'L-2', occupancyVal: 0, bpd: 8000 },
    'EK-TB': { age: 33, occupancy: 'L-2', occupancyVal: 0, bpd: 3000 }
};

// Traducir niveles de criticidad (Baja/Media/Alta) dinámicamente
function translateLevel(level, forceCase = 'none') {
    if (!level) return '';
    const upper = level.toString().toUpperCase().trim();
    let result = level;
    if (upper === 'BAJA') {
        result = currentLang === 'en' ? 'Low' : 'Baja';
    } else if (upper === 'MEDIA') {
        result = currentLang === 'en' ? 'Medium' : 'Media';
    } else if (upper === 'ALTA') {
        result = currentLang === 'en' ? 'High' : 'Alta';
    }
    
    if (forceCase === 'upper') return result.toUpperCase();
    if (forceCase === 'lower') return result.toLowerCase();
    if (forceCase === 'title') return result.charAt(0).toUpperCase() + result.slice(1).toLowerCase();
    return result;
}

// Traducir coordenadas de matriz (ej. ALTA-L-2 a HIGH-L-2) dinámicamente
function translateCoordinate(coord) {
    if (!coord) return '';
    const parts = coord.split('-');
    if (parts.length >= 2) {
        const lofPart = parts[0];
        const rest = parts.slice(1).join('-');
        return `${translateLevel(lofPart, 'upper')}-${rest}`;
    }
    return coord;
}

// Traducir tipos de estructura (Octápodo, Trípode, Tetrápodo, Monopilote) dinámicamente
function translateStructureType(tipo) {
    if (!tipo) return '';
    const upper = tipo.toUpperCase().trim();
    if (currentLang === 'en') {
        if (upper.includes('OCTÁPODO') || upper.includes('OCTAPODO')) return tipo.toUpperCase().includes('ACERO') ? 'Steel 8-Legged Platform' : '8-Legged Platform';
        if (upper.includes('TETRÁPODO') || upper.includes('TETRAPODO')) return tipo.toUpperCase().includes('ACERO') ? 'Steel 4-Legged Platform' : '4-Legged Platform';
        if (upper.includes('TRÍPODE ADOSADO') || upper.includes('TRIPODE ADOSADO')) return 'Attached Tripod Platform';
        if (upper.includes('TRÍPODE') || upper.includes('TRIPODE')) return 'Tripod Platform';
        if (upper.includes('MONOPILOTE')) return 'Monopile Platform';
    }
    return tipo;
}

// Obtener metadatos con fallbacks seguros
function getPlatformMeta(platName) {
    const meta = platformMetadata[platName] || {
        age: 15,
        occupancy: 'L-3',
        occupancyVal: 2,
        bpd: 2000
    };
    
    // Sobrescribir antigüedad con la cargada desde el Excel si está disponible
    if (platformsData[platName] && platformsData[platName]['ANTIGUEDAD']) {
        const loadedAge = parseInt(platformsData[platName]['ANTIGUEDAD'], 10);
        if (!isNaN(loadedAge)) {
            return { ...meta, age: loadedAge };
        }
    }
    return meta;
}


// Estado global de la aplicación
let currentLang = 'es';
let currentTab = 'dashboard';
let platformsData = {}; // Guarda las respuestas actuales (simuladas)
let userMitigations = {}; // Guarda los checkboxes de simulación por plataforma: { platName: { qId: true/false } }

// Inicialización de la aplicación
function mapOriginalToNew12(oldAnswers) {
    const newAnswers = {};
    
    // A.1: Régimen de Ocupación
    const oldA1 = oldAnswers['A1'] || 'a';
    if (oldA1 === 'c') newAnswers['A1'] = 'a';
    else if (oldA1 === 'b') newAnswers['A1'] = 'b';
    else newAnswers['A1'] = 'c';
    
    // A.2: Plan de Evacuación
    const oldA2 = oldAnswers['A2'] || 'a';
    if (newAnswers['A1'] === 'c') {
        newAnswers['A2'] = 'c';
    } else if (oldA2 === 'c') {
        newAnswers['A2'] = 'b';
    } else {
        newAnswers['A2'] = 'a';
    }
    
    // B.1: Riesgo por Fluido y H2S
    const oldA3 = oldAnswers['A3'] || 'a';
    if (oldA3 === 'c') newAnswers['B1'] = 'a';
    else if (oldA3 === 'b') newAnswers['B1'] = 'b';
    else newAnswers['B1'] = 'c';
    
    // B.2: Rol operativo
    const oldA4 = oldAnswers['A4'] || 'a';
    if (oldA4 === 'c') newAnswers['B2'] = 'a';
    else if (oldA4 === 'b') newAnswers['B2'] = 'b';
    else newAnswers['B2'] = 'c';
    
    // C.1: Air Gap
    const oldC1 = oldAnswers['C1'] || 'a';
    if (oldC1 === 'c') newAnswers['C1'] = 'a';
    else if (oldC1 === 'b') newAnswers['C1'] = 'b';
    else newAnswers['C1'] = 'c';
    
    // C.2: Redundancia
    const oldB4 = oldAnswers['B4'] || 'a';
    if (oldB4 === 'c') newAnswers['C2'] = 'a';
    else if (oldB4 === 'b') newAnswers['C2'] = 'b';
    else newAnswers['C2'] = 'c';
    
    // C.3: Estado de daño
    const oldB1 = oldAnswers['B1'] || 'a';
    const oldB2 = oldAnswers['B2'] || 'a';
    const oldB3 = oldAnswers['B3'] || 'a';
    const oldB5 = oldAnswers['B5'] || 'a';
    if (oldB1 === 'c' || oldB2 === 'c' || oldB3 === 'c' || oldB5 === 'c') {
        newAnswers['C3'] = 'a';
    } else if (oldB1 === 'b' || oldB2 === 'b' || oldB3 === 'b' || oldB5 === 'b') {
        newAnswers['C3'] = 'b';
    } else {
        newAnswers['C3'] = 'c';
    }
    
    // D.1: Personal
    const oldD1 = oldAnswers['D1'] || 'a';
    newAnswers['D1'] = (oldD1 === 'c' && (oldA1 === 'b' || oldA1 === 'c')) ? 'a' : 'b';
    
    // D.2: Instalaciones
    newAnswers['D2'] = (oldA3 === 'c') ? 'a' : 'b';
    
    // D.3: Sobrecarga
    newAnswers['D3'] = (oldC1 === 'c') ? 'a' : 'b';
    
    // D.4: Ola en Cubierta
    const oldD2 = oldAnswers['D2'] || 'a';
    newAnswers['D4'] = (oldD2 === 'c' || oldC1 === 'c') ? 'a' : 'b';
    
    // D.5: Daño Acumulado
    newAnswers['D5'] = (oldB1 === 'c' || oldB2 === 'c') ? 'a' : 'b';
    
    return newAnswers;
}

function ensure12QuestionsData(platAnswers) {
    if ('C3' in platAnswers && !('C4' in platAnswers)) {
        return platAnswers;
    }
    return mapOriginalToNew12(platAnswers);
}

document.addEventListener('DOMContentLoaded', () => {
    // Valores de metadatos por defecto para todas las plataformas (según datos del usuario)
    const defaultMetadata = {
        'EK-A Perf 2': { tipo: 'Octápodo', tirante: '51.3m', antiguedad: 12 },
        'Balam-1': { tipo: 'Sea Horse', tirante: '48.8m', antiguedad: 34 },
        'Balam-A': { tipo: 'Octápodo', tirante: '48.7m', antiguedad: 12 },
        'Balam-TA': { tipo: 'Trípode', tirante: '49.1m', antiguedad: 33 },
        'Balam-TB': { tipo: 'Tetrápodo', tirante: '48.8m', antiguedad: 33 },
        'Balam-TD': { tipo: 'Tetrápodo', tirante: '51.5m', antiguedad: 33 },
        'Balam-TE': { tipo: 'Tetrápodo', tirante: '51.5m', antiguedad: 33 },
        'EK-TA': { tipo: 'Trípode', tirante: '50.5m', antiguedad: 33 },
        'EK-TB': { tipo: 'Trípode', tirante: '49.9m', antiguedad: 33 },
        'Balam-TC': { tipo: 'Trípode', tirante: '52.0m', antiguedad: 33 },
        'EK-A Hab': { tipo: 'Octápodo', tirante: '51.0m', antiguedad: 17 },
        'EK-A': { tipo: 'Octápodo', tirante: '51.0m', antiguedad: 38 }
    };

    // Mapear base de datos inicial precargada del Excel y Módulo A directamente
    for (const [platName, sheetAnswers] of Object.entries(defaultExcelData)) {
        const answers = ensure12QuestionsData({ ...sheetAnswers });
        // Inyectar parámetros de diseño por defecto si existen
        if (defaultMetadata[platName]) {
            answers['TIPO'] = defaultMetadata[platName].tipo;
            answers['TIRANTE'] = defaultMetadata[platName].tirante;
            answers['ANTIGUEDAD'] = String(defaultMetadata[platName].antiguedad);
        }
        platformsData[platName] = answers;
    }

    // Cargar análisis activo previo de LocalStorage si existe
    const savedAnalysis = localStorage.getItem("simq_active_analysis");
    if (savedAnalysis) {
        try {
            const parsed = JSON.parse(savedAnalysis);
            if (parsed.platformsData) platformsData = parsed.platformsData;
            if (parsed.userMitigations) userMitigations = parsed.userMitigations;
        } catch(e) {
            console.error("Error loading active analysis from LocalStorage", e);
        }
    }

    initExcelUpload();
    translateStaticUI();
    checkSession();
});

// Cambiar de idioma
function setLanguage(lang) {
    currentLang = lang;
    
    // Sincronizar el select del header si existe
    const langSel = document.getElementById('langSelector');
    if (langSel && langSel.value !== lang) {
        langSel.value = lang;
    }

    translateStaticUI();
    switchTab(currentTab);
}

// Traducir elementos estáticos de la página
function translateStaticUI() {
    const t = uiTranslations[currentLang];
    
    // Elementos de la barra lateral (si existen en el DOM)
    const elSidebarSub = document.getElementById('lblSidebarSub');
    if (elSidebarSub) elSidebarSub.innerText = t.lblSidebarSub;
    
    const elActiveUser = document.getElementById('lblActiveUser');
    if (elActiveUser) elActiveUser.innerText = t.lblActiveUser;
    
    const elAnalysisMgmt = document.getElementById('lblAnalysisMgmt');
    if (elAnalysisMgmt) elAnalysisMgmt.innerText = t.lblAnalysisMgmt;
    
    const elLoadJsonBtn = document.getElementById('lblLoadJsonBtn');
    if (elLoadJsonBtn) elLoadJsonBtn.innerText = t.lblLoadJsonBtn;
    
    const elSaveAnalysis = document.getElementById('btnSaveAnalysis');
    if (elSaveAnalysis) elSaveAnalysis.innerHTML = t.btnSaveAnalysis;
    
    const elDownloadAnalysis = document.getElementById('btnDownloadAnalysis');
    if (elDownloadAnalysis) elDownloadAnalysis.innerHTML = t.btnDownloadAnalysis;
    
    const elExcelOps = document.getElementById('lblExcelOps');
    if (elExcelOps) elExcelOps.innerText = t.lblExcelOps;
    
    const elUploadZoneTitle = document.getElementById('uploadZoneTitle');
    if (elUploadZoneTitle) elUploadZoneTitle.innerText = t.uploadZoneTitle;
    
    const elUploadZoneSubtitle = document.getElementById('uploadZoneSubtitle');
    if (elUploadZoneSubtitle) elUploadZoneSubtitle.innerText = t.uploadZoneSubtitle;
    
    const elFooterSub = document.getElementById('lblFooterSub');
    if (elFooterSub) elFooterSub.innerText = t.lblFooterSub;
    
    // Encabezado
    const elPathApp = document.querySelector('.path-app');
    if (elPathApp) elPathApp.innerText = "SIM-Q RELIARISK";
    
    const elHeaderInst = document.getElementById('lblHeaderInst');
    if (elHeaderInst) elHeaderInst.innerText = currentLang === 'es' ? 'Área Contractual:' : 'Contract Area:';
    
    if (currentTab === 'dashboard') {
        const elCurrentPage = document.getElementById('headerCurrentPage');
        if (elCurrentPage) elCurrentPage.innerText = t.tabDashboard;
    } else if (currentTab.startsWith('platform:')) {
        const elCurrentPage = document.getElementById('headerCurrentPage');
        if (elCurrentPage) elCurrentPage.innerText = t.navSimulatorName;
    } else if (currentTab === 'admin') {
        const elCurrentPage = document.getElementById('headerCurrentPage');
        if (elCurrentPage) elCurrentPage.innerText = t.navAdminName;
    }
    
    const btnExportFleet = document.getElementById('btnExportFleet');
    if (btnExportFleet) {
        btnExportFleet.innerText = t.btnExportFleet;
    }
    
    // Botones de navegación de la barra lateral
    const elNavDashboard = document.getElementById('nav-dashboard');
    if (elNavDashboard) elNavDashboard.innerHTML = `📊 ${t.tabDashboard.replace('📊 ', '')}`;
    
    const elNavSimulator = document.getElementById('nav-simulator');
    if (elNavSimulator) elNavSimulator.innerHTML = `🎛️ ${t.navSimulatorName.replace('🎛️ ', '')}`;
    
    const elNavAdmin = document.getElementById('nav-admin');
    if (elNavAdmin) elNavAdmin.innerHTML = `👥 ${t.navAdminName.replace('👥 ', '')}`;
    
    const elLogoutBtn = document.querySelector('.logout-btn');
    if (elLogoutBtn) elLogoutBtn.innerHTML = `🚪 ${currentLang === 'es' ? 'Cerrar Sesión' : 'Logout'}`;
    
    // Campos del login (si está visible)
    const elLoginTitle = document.getElementById('loginTitle');
    if (elLoginTitle) elLoginTitle.innerText = "Reliarisk SIM-Q";
    
    const elLoginSubtitle = document.getElementById('loginSubtitle');
    if (elLoginSubtitle) elLoginSubtitle.innerText = currentLang === 'es' ? 'Sistema de Gestión de Integridad Estructural (API RP 2SIM)' : 'Structural Integrity Management System (API RP 2SIM)';
    
    const elLoginPrompt = document.getElementById('loginPrompt');
    if (elLoginPrompt) elLoginPrompt.innerText = currentLang === 'es' ? 'Ingrese sus credenciales de acceso para continuar:' : 'Enter your access credentials to continue:';
    
    const elLblUser = document.getElementById('lblUser');
    if (elLblUser) elLblUser.innerText = currentLang === 'es' ? 'Usuario / Client ID' : 'Username / Client ID';
    
    const elLblPass = document.getElementById('lblPass');
    if (elLblPass) elLblPass.innerText = currentLang === 'es' ? 'Contraseña / Password' : 'Password';
    
    const elBtnLogin = document.getElementById('btnLogin');
    if (elBtnLogin) elBtnLogin.innerText = currentLang === 'es' ? '🔐 Iniciar Sesión' : '🔐 Log In';
}

// Configurar el área de subida de Excel
function initExcelUpload() {
    const uploadZone = document.getElementById('uploadZoneSidebar');
    const excelFile = document.getElementById('excelFileSidebar');
    const statusDiv = document.getElementById('excelLoadStatus');

    if (!uploadZone || !excelFile) return;

    uploadZone.addEventListener('click', () => excelFile.click());

    excelFile.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            if (statusDiv) statusDiv.innerHTML = `<span style="color: var(--primary);">${currentLang === 'es' ? 'Procesando...' : 'Processing...'}</span>`;
            handleExcelFile(file);
        }
    });

    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });

    uploadZone.addEventListener('dragleave', () => {
        uploadZone.classList.remove('dragover');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) {
            if (statusDiv) statusDiv.innerHTML = `<span style="color: var(--primary);">${currentLang === 'es' ? 'Procesando...' : 'Processing...'}</span>`;
            handleExcelFile(file);
        }
    });
}

// Procesar el archivo de Excel usando SheetJS
function handleExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array', cellFormula: false});
            
            const newPlatformsData = {};
            let parsedCount = 0;

            workbook.SheetNames.forEach(sheetName => {
                const ws = workbook.Sheets[sheetName];
                const cleanName = sheetName.trim();
                
                const rows = XLSX.utils.sheet_to_json(ws, {header: 1});
                const sheetAnswers = {};
                let hasQuestions = false;
                
                const validQuestions = [
                    'A1','A2','B1','B2','C1','C2','C3','D1','D2','D3','D4','D5'
                ];
                const validMetadata = [
                    'TIPO', 'TIRANTE', 'ANTIGUEDAD'
                ];
                
                rows.forEach((row, rowIndex) => {
                    if (rowIndex === 0) return; // Omitir encabezado
                    if (!row || row.length < 3) return;
                    const qId = String(row[0]).trim().toUpperCase();
                    const val = String(row[2]).trim();
                    
                    if (validQuestions.includes(qId) && ['a', 'b', 'c'].includes(val.toLowerCase())) {
                        sheetAnswers[qId] = val.toLowerCase();
                        hasQuestions = true;
                    } else if (validMetadata.includes(qId)) {
                        sheetAnswers[qId] = val;
                    }
                });

                if (hasQuestions) {
                    newPlatformsData[cleanName] = sheetAnswers;
                    parsedCount++;
                }
            });

            if (parsedCount > 0) {
                platformsData = newPlatformsData;
                userMitigations = {};
                
                const statusDiv = document.getElementById('excelLoadStatus');
                if (statusDiv) {
                    statusDiv.innerHTML = `<span style="color: #10b981; font-weight: 600;">✅ ${currentLang === 'es' ? `Cargadas ${parsedCount} plataformas` : `Loaded ${parsedCount} platforms`}</span>`;
                } else {
                    alert(currentLang === 'es' ? `¡Se cargaron con éxito ${parsedCount} plataformas desde el archivo Excel!` : `Successfully loaded ${parsedCount} platforms from Excel file!`);
                }
                
                switchTab(currentTab);
            } else {
                if (currentLang === 'es') {
                    alert("No se encontraron hojas con el formato de cuestionario válido.");
                } else {
                    alert("No sheets with a valid questionnaire format were found.");
                }
            }

        } catch (error) {
            console.error(error);
            if (currentLang === 'es') {
                alert("Error al analizar el archivo de Excel. Asegúrate de subir el archivo correcto.");
            } else {
                alert("Error analyzing Excel file. Please make sure to upload the correct file.");
            }
        }
    };
    reader.readAsArrayBuffer(file);
}

let activeSimulatorPlatform = 'EK-A';

function buildTabBar() {
    // Deprecated in favor of left sidebar layout
}

// Cambiar de pestaña activa / enrutador
function switchTab(tabId) {
    if (!authenticated) return;
    
    // Quitar clase active de los botones de navegación de la barra lateral
    document.querySelectorAll('.sidebar-nav-btn').forEach(btn => btn.classList.remove('active'));
    
    const mainContent = document.getElementById('mainContent');
    mainContent.innerHTML = '';
    
    // Limpiar metadatos del encabezado
    const headerMetadata = document.getElementById('headerMetadataArea');
    headerMetadata.innerHTML = '';

    const isES = currentLang === 'es';
    const t = uiTranslations[currentLang];

    if (tabId === 'dashboard') {
        currentTab = 'dashboard';
        document.getElementById('nav-dashboard').classList.add('active');
        document.getElementById('headerCurrentPage').innerText = t.tabDashboard;
        
        // Mostrar botón de descarga en el encabezado
        const btnExportFleet = document.getElementById('btnExportFleet');
        if (btnExportFleet) btnExportFleet.style.display = 'inline-block';
        
        renderDashboard();
    } else if (tabId === 'simulator') {
        // Redirigir a la plataforma seleccionada activa
        if (!platformsData[activeSimulatorPlatform]) {
            activeSimulatorPlatform = Object.keys(platformsData)[0] || 'EK-A';
        }
        switchTab('platform:' + activeSimulatorPlatform);
    } else if (tabId.startsWith('platform:')) {
        const platName = tabId.substring(9);
        currentTab = tabId;
        activeSimulatorPlatform = platName;
        
        document.getElementById('nav-simulator').classList.add('active');
        document.getElementById('headerCurrentPage').innerText = t.navSimulatorName;
        
        // Ocultar reporte de flota (cada plataforma tiene su propio botón de exportar)
        const btnExportFleet = document.getElementById('btnExportFleet');
        if (btnExportFleet) btnExportFleet.style.display = 'none';

        // Agregar metadatos de plataforma al encabezado
        const meta = getPlatformMeta(platName);
        const answers = platformsData[platName] || {};
        const structureTypeRaw = (answers && answers['TIPO']) ? answers['TIPO'] 
                            : (platName.includes('Hab') ? 'Tetrápodo de Acero'
                            : (platName.includes('TA') && !platName.includes('Balam') ? 'Trípode Adosado'
                            : (platName.includes('TB') ? 'Trípode Adosado'
                            : (platName.includes('TC') ? 'Trípode'
                            : (platName.includes('1') ? (isES ? 'Estructura Marina Ligera' : 'Light Sea Horse Structure') : 'Octápodo de Acero')))));
        const structureType = isES ? structureTypeRaw : translateStructureType(structureTypeRaw);
        const waterDepth = (answers && answers['TIRANTE']) 
                            ? (String(answers['TIRANTE']).endsWith('m') ? answers['TIRANTE'] : `${answers['TIRANTE']}m`)
                            : (platName.includes('TA') ? '51.5m' : (platName.includes('TB') ? '51.5m' : (platName.includes('A') ? '50.0m' : '52.0m')));
        
        const ageText = isES ? `${meta.age} años` : `${meta.age} years`;
        headerMetadata.innerHTML = `
            <div class="metadata-item">
                <span class="metadata-label">${isES ? 'Plataforma' : 'Platform'}:</span>
                <span class="metadata-val" style="color: #00daf3; font-weight: 700;">${platName}</span>
            </div>
            <div class="metadata-item">
                <span class="metadata-label">${isES ? 'Tipo de Estructura' : 'Structure Type'}:</span>
                <span class="metadata-val">${structureType}</span>
            </div>
            <div class="metadata-item">
                <span class="metadata-label">${isES ? 'Tirante' : 'Water Depth'}:</span>
                <span class="metadata-val">${waterDepth}</span>
            </div>
            <div class="metadata-item">
                <span class="metadata-label">${isES ? 'Antigüedad' : 'Age'}:</span>
                <span class="metadata-val">${ageText}</span>
            </div>
        `;
        
        // Crear el selector de plataforma en la parte superior
        const selectorCard = document.createElement('div');
        selectorCard.className = 'simulator-view-header';
        
        const platformOptions = Object.keys(platformsData).map(name => {
            const isSelected = name === platName ? 'selected' : '';
            return `<option value="${name}" ${isSelected}>${name}</option>`;
        }).join('');
        
        selectorCard.innerHTML = `
            <div class="simulator-selector-card">
                <label for="platformSelect">${t.lblSelectPlatform}</label>
                <select id="platformSelect" onchange="switchTab('platform:' + this.value)">
                    ${platformOptions}
                </select>
            </div>
        `;
        mainContent.appendChild(selectorCard);

        renderPlatform(platName);
    } else if (tabId === 'admin') {
        currentTab = 'admin';
        document.getElementById('nav-admin').classList.add('active');
        document.getElementById('headerCurrentPage').innerText = t.navAdminName;
        
        const btnExportFleet = document.getElementById('btnExportFleet');
        if (btnExportFleet) btnExportFleet.style.display = 'none';
        
        renderAdminPanel();
    }
}

// Calcular ejes para el gráfico radial
function calculateRadarAxes(answers) {
    // Eje 1: Luz Libre / Air Gap (C1)
    const c1 = answers['C1'] || 'c';
    const airGapVal = c1 === 'c' ? 5 : (c1 === 'b' ? 3 : 1);
    
    // Eje 2: Redundancia Jacket (C2)
    const c2 = answers['C2'] || 'c';
    const redVal = c2 === 'c' ? 5 : (c2 === 'b' ? 3 : 1);
    
    // Eje 3: Daño Estructural (C3)
    const c3 = answers['C3'] || 'c';
    const dmgVal = c3 === 'c' ? 5 : (c3 === 'b' ? 3 : 1);
    
    // Eje 4: Impacto y Fluidos (B1, B2)
    const b1 = answers['B1'] || 'c';
    const b2 = answers['B2'] || 'c';
    let consVal = 5;
    if (b1 === 'a' || b2 === 'a') consVal = 1;
    else if (b1 === 'b' || b2 === 'b') consVal = 3;
    
    // Eje 5: Iniciadores de Reevaluación (D1-D5)
    let activeCount = 0;
    ['D1', 'D2', 'D3', 'D4', 'D5'].forEach(qId => {
        if (answers[qId] === 'a') activeCount++;
    });
    const initVal = Math.max(1, 5 - activeCount);
    
    return [airGapVal, redVal, dmgVal, consVal, initVal];
}

// Calcular datos para el gráfico de cascada
function calculateWaterfallData(answers) {
    const data = [];
    const questionsOrder = ['C1', 'C2', 'C3', 'D1', 'D2', 'D3', 'D4', 'D5'];
    
    questionsOrder.forEach(qId => {
        const ans = answers[qId] || 'c';
        let pts = 0;
        if (qId.startsWith('C')) {
            if (ans === 'a') pts = 3;
            else if (ans === 'b') pts = 1;
        } else { // Module D
            if (ans === 'a') pts = 2;
        }
        
        if (pts > 0) {
            const qMeta = questionsMetadata.find(q => q.id === qId);
            data.push({
                id: qId,
                label: qId,
                title: qMeta ? qMeta.label[currentLang] : qId,
                value: pts
            });
        }
    });
    return data;
}

// Calcular las métricas completas para una plataforma (API RP 2SIM SIM Engine)
function calculatePlatformMetrics(platName, answers) {
    // 1. Determinar Categoría de Ocupación Humana (S-1, S-2, S-3)
    const a1 = answers['A1'] || 'c';
    const a2 = answers['A2'] || 'c';
    let sCategory = 'S-3';
    if (a1 === 'c' || a2 === 'c') {
        sCategory = 'S-3';
    } else if (a2 === 'b') {
        sCategory = 'S-1';
    } else if ((a1 === 'a' || a1 === 'b') && a2 === 'a') {
        sCategory = 'S-2';
    }
    
    // Special override for quarters platform to be S-1 (Manned-nonevacuated) as per document
    if (platName === 'EK-A-Hab' || platName === 'EK-A Hab') {
        sCategory = 'S-1';
    }

    // 2. Determinar Categoría de Consecuencia (C-1, C-2, C-3)
    const b1 = answers['B1'] || 'c';
    const b2 = answers['B2'] || 'c';
    let cCategory = 'C-3';
    if (b1 === 'a' || b2 === 'a') {
        cCategory = 'C-1';
    } else if (b1 === 'b' || b2 === 'b') {
        cCategory = 'C-2';
    } else {
        cCategory = 'C-3';
    }

    // 3. Determinar Categoría de Exposición (L-1, L-2, L-3)
    let exposure = 'L-3';
    if (sCategory === 'S-1') {
        exposure = 'L-1';
    } else if (sCategory === 'S-2') {
        if (cCategory === 'C-1') exposure = 'L-1';
        else exposure = 'L-2';
    } else if (sCategory === 'S-3') {
        if (cCategory === 'C-1') exposure = 'L-1';
        else if (cCategory === 'C-2') exposure = 'L-2';
        else exposure = 'L-3';
    }

    // 4. Determinar Probabilidad de Falla (LoF: BAJA, MEDIA, ALTA)
    const c1 = answers['C1'] || 'c';
    const c2 = answers['C2'] || 'c';
    const c3 = answers['C3'] || 'c';
    
    let lof = 'MEDIA';
    if (c1 === 'a' || c3 === 'a') {
        lof = 'ALTA';
    } else if (c1 === 'c' && c2 === 'c' && c3 === 'c') {
        lof = 'BAJA';
    } else {
        lof = 'MEDIA';
    }

    // 5. Cruzamiento de Riesgo
    const riskCode = matrixValues[lof][exposure] || 'L';
    const risk = riskLevels[riskCode] || riskLevels['L'];

    // 6. Evaluación de Iniciadores (Módulo D)
    let hasInitiatorActive = false;
    ['D1', 'D2', 'D3', 'D4', 'D5'].forEach(qId => {
        if (answers[qId] === 'a') {
            hasInitiatorActive = true;
        }
    });

    // Sum points for condition & initiators for waterfall graph
    let lofPoints = 0;
    if (c1 === 'a') lofPoints += 3; else if (c1 === 'b') lofPoints += 1;
    if (c2 === 'a') lofPoints += 3; else if (c2 === 'b') lofPoints += 1;
    if (c3 === 'a') lofPoints += 3; else if (c3 === 'b') lofPoints += 1;
    ['D1', 'D2', 'D3', 'D4', 'D5'].forEach(qId => {
        if (answers[qId] === 'a') lofPoints += 2;
    });

    return {
        probPoints: lofPoints,
        probLevel: lof, // 'BAJA', 'MEDIA', 'ALTA'
        finalCons: exposure, // 'L-1', 'L-2', 'L-3'
        riskCode: riskCode, // 'L', 'M', 'H'
        riskName: risk.name[currentLang],
        riskClass: risk.class,
        riskColor: risk.color,
        coordinate: `${lof}-${exposure}`,
        triggerEscalation: hasInitiatorActive,
        radarAxes: calculateRadarAxes(answers),
        waterfallData: calculateWaterfallData(answers)
    };
}

// Obtener intervalos de inspección recomendados (GVI / CVI / NDT) según riesgo
function getInspectionIntervals(riskCode) {
    if (riskCode === 'H') {
        return {
            gvi: { es: "1 - 2 años", en: "1 - 2 years" },
            cvi: { es: "3 - 5 años (Dirigida)", en: "3 - 5 years (Directed)" },
            ndt: { es: "3 - 5 años (ACFM/FMD)", en: "3 - 5 years (ACFM/FMD)" }
        };
    }
    if (riskCode === 'M') {
        return {
            gvi: { es: "3 - 5 años", en: "3 - 5 years" },
            cvi: { es: "6 - 10 años", en: "6 - 10 years" },
            ndt: { es: "Subacuático (Monitoreo)", en: "Subsea (Monitoring)" }
        };
    }
    return {
        gvi: { es: "5 - 6 años", en: "5 - 6 years" },
        cvi: { es: "10 - 12 años", en: "10 - 12 years" },
        ndt: { es: "Según campaña regular", en: "As per regular campaign" }
    };
}

// RENDERIZAR DASHBOARD GENERAL
function renderDashboard() {
    const mainContent = document.getElementById('mainContent');
    const template = document.getElementById('dashboard-template');
    const clone = template.content.cloneNode(true);
    mainContent.appendChild(clone);

    const t = uiTranslations[currentLang];
    document.querySelector('.matrix-panel h2').innerText = t.matrixTitle;
    document.querySelector('.matrix-panel .panel-subtitle').innerText = t.matrixSubtitle;
    document.querySelector('.table-panel h2').innerText = t.tableTitle;
    document.querySelector('.table-panel .panel-subtitle').innerText = t.tableSubtitle;

    const ths = document.querySelectorAll('.hierarchy-table th');
    ths[0].innerText = t.colPlat;
    ths[1].innerText = t.colBaseCons;
    ths[2].innerText = t.colPoints;
    ths[3].innerText = t.colCoord;
    ths[4].innerText = t.colRisk;
    ths[5].innerText = t.colAction;

    const platforms = [];
    let countTotal = 0;
    let countHigh = 0;
    let countMedium = 0;
    let countLow = 0;

    Object.keys(platformsData).forEach(platName => {
        const metrics = calculatePlatformMetrics(platName, platformsData[platName]);
        platforms.push({ name: platName, metrics: metrics });

        countTotal++;
        if (metrics.riskCode === 'H') countHigh++;
        else if (metrics.riskCode === 'M') countMedium++;
        else if (metrics.riskCode === 'L') countLow++;
    });

    document.getElementById('kpiTotal').innerText = countTotal;
    document.getElementById('kpiHigh').innerText = countHigh;
    document.getElementById('kpiMedium').innerText = countMedium;
    document.getElementById('kpiLow').innerText = countLow;
    
    const kpiCards = document.querySelectorAll('.kpi-card');
    kpiCards[0].querySelector('.kpi-lbl').innerText = t.kpiTotal;
    kpiCards[1].querySelector('.kpi-lbl').innerText = t.kpiHigh;
    kpiCards[2].querySelector('.kpi-lbl').innerText = t.kpiMedium;
    kpiCards[3].querySelector('.kpi-lbl').innerText = t.kpiLow;

    renderGlobalMatrix(platforms);

    const riskSortWeight = { 'H': 3, 'M': 2, 'L': 1 };
    platforms.sort((a, b) => {
        const weightA = riskSortWeight[a.metrics.riskCode] || 0;
        const weightB = riskSortWeight[b.metrics.riskCode] || 0;
        if (weightA !== weightB) return weightB - weightA;
        
        // Secondary sort by vulnerability score
        return b.metrics.probPoints - a.metrics.probPoints;
    });

    const tbody = document.getElementById('hierarchyTableBody');
    tbody.innerHTML = '';
    
    platforms.forEach(plat => {
        const tr = document.createElement('tr');
        const m = plat.metrics;
        
        let riskClassClean = m.riskClass;
        let riskNameClean = m.riskName;
        let triggerMarker = m.triggerEscalation ? ' ⚠️' : '';

        let inactiveBadge = '';
        if (plat.name === 'Balam-TC') {
            inactiveBadge = ` <span class="badge-risk" style="background-color: #64748b; color: white; margin-left: 5px; font-size: 9px; padding: 2px 6px;">${currentLang === 'es' ? 'INACTIVA' : 'INACTIVE'}</span>`;
        }

        tr.innerHTML = `
            <td>
                <strong>${plat.name}</strong>${inactiveBadge}
                <div class="plat-desc-sub">${platformDescriptions[plat.name] ? platformDescriptions[plat.name][currentLang] : t.platDescDefault}</div>
            </td>
            <td>${m.finalCons}</td>
            <td>${translateLevel(m.probLevel, 'upper')}</td>
            <td><span class="matrix-coord">${translateCoordinate(m.coordinate)}</span></td>
            <td><span class="badge-risk ${riskClassClean}">${riskNameClean}${triggerMarker}</span></td>
            <td><button class="action-btn" onclick="switchTab('platform:${plat.name}')">${t.btnInspect}</button></td>
        `;
        tbody.appendChild(tr);
    });

    // Traducir títulos de la sección de Analíticas de Flota
    document.getElementById('lblAnalyticsTitle').innerText = t.lblAnalyticsTitle;
    document.getElementById('lblAnalyticsSubtitle').innerText = t.lblAnalyticsSubtitle;
    document.getElementById('btn-tab-bubble').innerText = t.tabBubble;
    document.getElementById('btn-tab-pareto').innerText = t.tabPareto;
    document.getElementById('btn-tab-sankey').innerText = t.tabSankey;

    switchAnalyticsTab(currentAnalyticsTab);
}

// Renderizar Matriz Global (3x5)
function renderGlobalMatrix(platforms) {
    const grid = document.getElementById('matrixGridGlobal');
    grid.innerHTML = '';

    const lofOrder = ['BAJA', 'MEDIA', 'ALTA'];
    const cofOrder = ['L-1', 'L-2', 'L-3'];
    const t = uiTranslations[currentLang];

    grid.innerHTML += `<div class="matrix-cell-global corner-cell">
        <div class="lbl-lof">LoF</div>
        <div class="lbl-divider"></div>
        <div class="lbl-cof">CoF</div>
    </div>`;
    
    lofOrder.forEach(l => {
        grid.innerHTML += `<div class="matrix-cell-global header-cell">${translateLevel(l, 'upper')}</div>`;
    });

    cofOrder.forEach(c => {
        grid.innerHTML += `<div class="matrix-cell-global header-cell">${c}</div>`;
        
        lofOrder.forEach(l => {
            const riskCode = matrixValues[l][c];
            const risk = riskLevels[riskCode];
            
            const matching = platforms.filter(plat => plat.metrics.probLevel === l && plat.metrics.finalCons === c);
            
            let badgesHtml = '';
            matching.forEach(plat => {
                const triggerAlert = plat.metrics.triggerEscalation ? 'style="border: 2px solid #ef4444;"' : '';
                badgesHtml += `
                    <button class="platform-badge-link" 
                       ${triggerAlert}
                       onclick="switchTab('platform:${plat.name}')" 
                       title="${plat.name} (${t.colRisk}: ${plat.metrics.riskName})">
                       ${plat.name}
                    </button>
                `;
            });

            grid.innerHTML += `
                <div class="matrix-cell-global data-cell ${risk.class}">
                    <div class="badges-container">${badgesHtml}</div>
                </div>
            `;
        });
    });
}

// RENDERIZAR DETALLE DE PLATAFORMA INDIVIDUAL
function renderPlatform(platName) {
    const mainContent = document.getElementById('mainContent');
    const template = document.getElementById('platform-template');
    const clone = template.content.cloneNode(true);

    if (!userMitigations[platName]) {
        userMitigations[platName] = {};
    }

    const t = uiTranslations[currentLang];
    
    // Configurar título y descripción directamente en el fragmento clonado
    clone.querySelector('#platformNameHeader').innerText = platName;
    clone.querySelector('#platformDescriptionText').innerText = platformDescriptions[platName] ? platformDescriptions[platName][currentLang] : t.platDescDefault;
    
    // Banner para plataforma inactiva (Balam-TC)
    if (platName === 'Balam-TC') {
        const headerRow = clone.querySelector('.platform-header-row');
        if (headerRow) {
            const banner = document.createElement('div');
            banner.className = 'inactive-platform-banner';
            banner.style.width = '100%';
            banner.style.gridColumn = '1 / -1';
            banner.style.marginTop = '10px';
            banner.style.padding = '12px 16px';
            banner.style.borderRadius = '8px';
            banner.style.background = 'var(--surface-container-low)';
            banner.style.outline = '1px solid var(--outline-variant)';
            banner.style.color = 'var(--text-muted)';
            banner.style.fontSize = '0.85rem';
            banner.style.display = 'flex';
            banner.style.gap = '10px';
            banner.style.alignItems = 'center';
            banner.innerHTML = `⚠️ <span>${currentLang === 'es' 
                ? '<strong>Atención:</strong> Esta plataforma se encuentra actualmente <strong>fuera de operación / inactiva</strong>. Los parámetros mostrados reflejan su condición inactiva.' 
                : '<strong>Attention:</strong> This platform is currently <strong>out of operation / inactive</strong>. The parameters displayed reflect its inactive condition.'}</span>`;
            
            const layout = clone.querySelector('.platform-layout');
            if (layout) {
                layout.insertBefore(banner, layout.children[1]);
            }
        }
    }
    
    // Traducir elementos directamente en el fragmento clonado antes de insertarlo en el DOM
    clone.querySelector('.form-section h2').innerText = t.platTitleQuestions;
    clone.querySelector('.form-section .panel-subtitle').innerText = t.platSubtitleQuestions;
    clone.querySelector('.matrix-container h2').innerText = t.platTitleMatrix;
    clone.querySelector('.summary-container h2').innerText = t.platTitleStats;
    clone.querySelector('.mitigation-panel h2').innerText = t.platTitleMitigation;
    clone.querySelector('.mitigation-panel .panel-subtitle').innerHTML = t.platSubtitleMitigation;
    clone.querySelector('#inspectionCardTitle').innerText = t.inspectionTitle;

    clone.querySelector('.platform-charts-panel h2').innerText = t.platformChartsTitle;
    clone.querySelector('.platform-charts-panel .panel-subtitle').innerText = t.platformChartsSubtitle;
    clone.querySelector('#lblRadarChart').innerText = t.lblRadarChart;
    clone.querySelector('#lblWaterfallChart').innerText = t.lblWaterfallChart;

    const btnExportPlatform = clone.querySelector('#btnExportPlatform');
    if (btnExportPlatform) {
        btnExportPlatform.innerText = t.btnExportPlatform;
    }

    // Insertar el fragmento clonado ya configurado en el documento activo
    mainContent.appendChild(clone);

    // Renderizar preguntas modularizadas (incluyendo Módulo A)
    renderPlatformQuestions(platName);

    // Actualizar simulaciones y dibujar resultados
    updatePlatformSimulation();
}


// Renderizar las preguntas en el cuestionario individual
function renderPlatformQuestions(platName) {
    const container = document.getElementById('questionList');
    container.innerHTML = '';

    const currentAnswers = platformsData[platName];
    const mitigations = userMitigations[platName];
    const t = uiTranslations[currentLang];

    let currentModule = '';

    questionsMetadata.forEach(q => {
        if (q.module !== currentModule) {
            currentModule = q.module;
            const modHeader = document.createElement('div');
            modHeader.className = 'module-separator-header';
            modHeader.innerText = q.moduleName[currentLang];
            container.appendChild(modHeader);
        }

        const qCard = document.createElement('div');
        qCard.className = 'question-card';
        qCard.id = `qcard-${q.id}`;

        const ans = currentAnswers[q.id] || 'c';
        const isSimulated = mitigations[q.id] === true;
        
        // Define default simulated value: safest is 'c' for modules A-C, and 'b' (No) for module D.
        const defaultSafestAns = q.module === 'D' ? 'b' : 'c';
        const activeAns = isSimulated ? defaultSafestAns : ans;

        const defaultComments = defaultExcelData[platName] && defaultExcelData[platName][q.id] 
            ? defaultExcelData[platName][q.id].comments 
            : '';

        let optionsHtml = `
            <label class="radio-label">
                <input type="radio" name="radio-${q.id}" value="a" ${activeAns === 'a' ? 'checked' : ''} 
                    ${isSimulated ? 'disabled' : ''} onchange="handleAnswerChange('${platName}', '${q.id}', 'a')">
                <span><strong>a)</strong> ${q.options.a[currentLang]}</span>
            </label>
            <label class="radio-label">
                <input type="radio" name="radio-${q.id}" value="b" ${activeAns === 'b' ? 'checked' : ''} 
                    ${isSimulated ? 'disabled' : ''} onchange="handleAnswerChange('${platName}', '${q.id}', 'b')">
                <span><strong>b)</strong> ${q.options.b[currentLang]}</span>
            </label>
        `;
        
        if (q.options.c) {
            optionsHtml += `
                <label class="radio-label">
                    <input type="radio" name="radio-${q.id}" value="c" ${activeAns === 'c' ? 'checked' : ''} 
                        ${isSimulated ? 'disabled' : ''} onchange="handleAnswerChange('${platName}', '${q.id}', 'c')">
                    <span class="${q.isPrincipal ? 'text-critical' : ''}"><strong>c)</strong> ${q.options.c[currentLang]}</span>
                </label>
            `;
        }

        qCard.innerHTML = `
            <div class="question-header">
                <span class="question-code">${q.id}</span>
                <span class="question-text">${q.text[currentLang]}</span>
            </div>
            <div class="radio-group-vertical">
                ${optionsHtml}
            </div>
            ${defaultComments ? `<div class="question-comment"><strong>Obs:</strong> ${defaultComments}</div>` : ''}
        `;
        container.appendChild(qCard);
    });
}

// Guardar cambio manual de respuesta
function handleAnswerChange(platName, qId, value) {
    platformsData[platName][qId] = value;
    updatePlatformSimulation();
}

// Calcular y actualizar simulación en tiempo real
function updatePlatformSimulation() {
    const platName = currentTab.substring(9);
    if (!platName) return;

    const mitigations = userMitigations[platName];
    const originalAnswers = platformsData[platName];

    // Construir conjunto de respuestas simuladas (mitigado -> valor más seguro: 'c' para A-C, 'b' para D)
    const simulatedAnswers = {};
    Object.keys(originalAnswers).forEach(qId => {
        if (mitigations[qId] === true) {
            simulatedAnswers[qId] = qId.startsWith('D') ? 'b' : 'c';
        } else {
            simulatedAnswers[qId] = originalAnswers[qId];
        }
    });

    const origMetrics = calculatePlatformMetrics(platName, originalAnswers);
    const simMetrics = calculatePlatformMetrics(platName, simulatedAnswers);

    // Actualizar Badges superiores
    const riskBadge = document.getElementById('platformRiskBadge');
    riskBadge.innerText = simMetrics.riskName;
    riskBadge.className = `risk-badge-large ${simMetrics.riskClass}`;
    
    const coordBadge = document.getElementById('platformCoordBadge');
    coordBadge.innerText = translateCoordinate(simMetrics.coordinate);

    const triggerBadge = document.getElementById('triggerWarningBadge');
    if (simMetrics.triggerEscalation) {
        triggerBadge.style.display = 'inline-flex';
        triggerBadge.innerText = uiTranslations[currentLang].triggerBadge;
    } else {
        triggerBadge.style.display = 'none';
    }

    // Actualizar caja de estadísticas
    document.getElementById('probLevel').innerText = translateLevel(simMetrics.probLevel, 'title');
    document.getElementById('consFinal').innerText = simMetrics.finalCons;
    document.getElementById('riskLevelVal').innerText = simMetrics.riskName;

    const t = uiTranslations[currentLang];
    const statBoxes = document.querySelectorAll('.stat-box');
    statBoxes[0].querySelector('.stat-label').innerText = t.statProbLevel;
    statBoxes[1].querySelector('.stat-label').innerText = t.statConsFinal;
    statBoxes[2].querySelector('.stat-label').innerText = t.colRisk;

    // Dibujar la matriz 3x3 de simulación
    renderPlatformMatrix(origMetrics.coordinate, simMetrics.coordinate);

    // Generar Plan de Mitigación y Recomendaciones
    renderMitigationPlan(platName, originalAnswers, mitigations, simMetrics);

    // Dibujar gráficos individuales
    drawRadarChart('radar-chart-container', simMetrics.radarAxes);
    drawWaterfallChart('waterfall-chart-container', simMetrics.waterfallData, simMetrics.probPoints);
}

// Renderizar Matriz Individual de Plataforma (3x5)
function renderPlatformMatrix(originalCoord, simulatedCoord) {
    const grid = document.getElementById('matrixGrid');
    grid.innerHTML = '';

    const lofOrder = ['BAJA', 'MEDIA', 'ALTA'];
    const cofOrder = ['L-1', 'L-2', 'L-3'];
    const t = uiTranslations[currentLang];

    grid.innerHTML += `<div class="matrix-cell corner-cell">
        <div class="lbl-lof">LoF</div>
        <div class="lbl-divider"></div>
        <div class="lbl-cof">CoF</div>
    </div>`;

    lofOrder.forEach(l => {
        grid.innerHTML += `<div class="matrix-cell header-cell">${translateLevel(l, 'upper')}</div>`;
    });

    cofOrder.forEach(c => {
        grid.innerHTML += `<div class="matrix-cell header-cell">${c}</div>`;
        lofOrder.forEach(l => {
            const riskCode = matrixValues[l][c];
            const risk = riskLevels[riskCode];
            const coord = `${l}-${c}`;
            
            let cellClass = `matrix-cell ${risk.class}`;
            let label = risk.name[currentLang];
            
            if (coord === simulatedCoord) {
                cellClass += ' active';
            }

            let marker = '';
            if (coord === originalCoord && coord === simulatedCoord) {
                marker = `<span class="cell-marker original-simulated">${t.originalLabel} & Sim</span>`;
            } else if (coord === originalCoord) {
                marker = `<span class="cell-marker original">${t.originalLabel}</span>`;
            } else if (coord === simulatedCoord) {
                marker = `<span class="cell-marker simulated">Sim</span>`;
            }

            grid.innerHTML += `
                <div class="${cellClass}">
                    ${marker}
                </div>
            `;
        });
    });
}

// Obtener acción de mitigación personalizada para plataformas inactivas (ej. Balam-TC)
function getMitigationAction(platName, qId, q, currentLang) {
    if (platName === 'Balam-TC') {
        if (qId === 'C2') {
            return currentLang === 'es' 
                ? 'Para esta estructura inactiva (L-3), postergar el refuerzo estructural físico de nudos críticos e instalar en su lugar un programa de monitoreo de estabilidad global pasiva.' 
                : 'For this inactive structure (L-3), defer physical reinforcement of critical nodes and implement a passive global stability monitoring program.';
        }
        if (qId === 'C3') {
            return currentLang === 'es' 
                ? 'Monitorear de forma no intervencionista la degradación del Jacket mediante inspección visual y mediciones de potencial catódico; posponer soldaduras o reparaciones subacuáticas mayores salvo que comprometa ductos o colectores activos adyacentes.' 
                : 'Perform non-interventionist monitoring of jacket degradation via visual inspection and cathodic potential measurements; postpone major subsea repairs or welding unless adjacent active pipelines are threatened.';
        }
        if (qId === 'D4') {
            return currentLang === 'es' 
                ? 'Exentar de análisis de re-calificación Pushover Nivel III detallado por su condición inactiva y nula exposición a riesgos humanos o ambientales mayores; mantener monitoreo visual de marcas de impacto de ola en cubierta inferior.' 
                : 'Exempt from detailed Level III Pushover re-qualification analysis due to its inactive condition and zero exposure to major safety or environmental risks; maintain visual monitoring of wave impact marks on cellar deck.';
        }
    }
    return q.mitigationAction[currentLang];
}

// Renderizar panel de mitigación y recomendaciones de intervalo de inspección
function renderMitigationPlan(platName, originalAnswers, mitigations, sim) {
    const listContainer = document.getElementById('mitigationActionsList');
    const banner = document.getElementById('mitigationStatusBanner');
    const inspectionCard = document.getElementById('inspectionRecommendationCard');
    
    listContainer.innerHTML = '';
    const t = uiTranslations[currentLang];

    const intervals = getInspectionIntervals(sim.riskCode);
    document.getElementById('gviInterval').innerText = intervals.gvi[currentLang];
    document.getElementById('cviInterval').innerText = intervals.cvi[currentLang];
    document.getElementById('ndtInterval').innerText = intervals.ndt[currentLang];
    
    const intervalLbls = document.querySelectorAll('.interval-lbl');
    intervalLbls[0].innerText = t.gviLabel;
    intervalLbls[1].innerText = t.cviLabel;
    intervalLbls[2].innerText = t.ndtLabel;
    inspectionCard.style.display = 'block';

    if (sim.riskCode === 'L') {
        const hasSimulations = Object.values(mitigations).some(v => v === true);
        if (platName === 'Balam-TC') {
            if (!hasSimulations) {
                banner.className = "mitigation-status-banner info";
                banner.innerHTML = currentLang === 'es'
                    ? "📋 <strong>Estatus Especial - Plataforma Inactiva con Deficiencias:</strong> Aunque la plataforma se encuentra en nivel de <strong>Riesgo Bajo (L)</strong> debido a su condición de inactividad (0 POB y 0 BPD), se identificaron deficiencias estructurales e iniciadores de reevaluación. Se recomienda un <strong>Plan de Mitigación Pasivo y Monitoreo</strong> que postergue reparaciones estructurales mayores o análisis Pushover detallados en favor de un seguimiento visual y de estabilidad global."
                    : "📋 <strong>Special Status - Inactive Platform with Deficiencies:</strong> Although the platform is at <strong>Low Risk (L)</strong> due to its inactive condition (0 POB & 0 BPD), structural deficiencies and evaluation initiators were identified. A <strong>Passive Mitigation & Monitoring Plan</strong> is recommended, deferring major structural repairs or detailed Pushover analyses in favor of visual and global stability monitoring.";
            } else {
                banner.className = "mitigation-status-banner success";
                banner.innerHTML = currentLang === 'es'
                    ? "✔️ <strong>Mitigación Simulada para Plataforma Inactiva:</strong> Se ha simulado la mitigación de las deficiencias seleccionadas. Por su condición inactiva, la plataforma mantiene su clasificación de <strong>Riesgo Bajo (L)</strong>."
                    : "✔️ <strong>Simulated Mitigation for Inactive Platform:</strong> Deficiency mitigation has been simulated. Due to its inactive condition, the platform maintains its <strong>Low Risk (L)</strong> classification.";
            }
        } else if (!hasSimulations) {
            banner.className = "mitigation-status-banner optimal";
            banner.innerHTML = t.successOptimal;
        } else {
            banner.className = "mitigation-status-banner success";
            banner.innerHTML = t.successGoal.replace('{coord}', translateCoordinate(sim.coordinate));
        }
    } else {
        banner.className = "mitigation-status-banner warning";
        
        let reqText = '';
        if (sim.finalCons === 'L-1') {
            reqText = t.reqTextCons1;
        } else if (sim.finalCons === 'L-2') {
            reqText = t.reqTextCons2;
        } else {
            reqText = t.reqTextCons3;
        }

        banner.innerHTML = t.bannerWarning
            .replace('{riskName}', sim.riskName)
            .replace('{coord}', translateCoordinate(sim.coordinate))
            .replace('{reqText}', reqText);
    }

    let hasFindings = false;

    questionsMetadata.forEach(q => {
        const origAns = originalAnswers[q.id] || 'c';
        
        // A finding exists if the answer is not the safest ('c' for A-C, 'a' for D)
        let isFinding = false;
        if (q.module === 'D') {
            isFinding = (origAns === 'a');
        } else {
            isFinding = (origAns === 'a' || origAns === 'b');
        }

        if (isFinding) {
            hasFindings = true;
            
            const isMitigated = mitigations[q.id] === true;
            const card = document.createElement('div');
            card.className = `mitigation-card ${isMitigated ? 'completed' : ''}`;
            
            let ptsDeduction = 0;
            if (q.id.startsWith('C')) {
                ptsDeduction = origAns === 'a' ? 3 : 1;
            } else if (q.id.startsWith('D')) {
                ptsDeduction = 2;
            }

            const deductionLabel = q.type === 'consequence' 
                ? (currentLang === 'es' ? 'Mitigación CoF' : 'CoF Mitigation')
                : `-${ptsDeduction} pts LoF`;

            card.innerHTML = `
                <div class="mitigation-checkbox-wrapper" onclick="toggleMitigation('${platName}', '${q.id}')">
                    <div class="mitigation-checkbox"></div>
                </div>
                <div class="mitigation-card-content">
                    <div class="mitigation-card-header">
                        <span class="mitigation-title">${q.label[currentLang]}</span>
                        <span class="mitigation-points-deduction" style="${q.type === 'consequence' ? 'background: rgba(16, 185, 129, 0.15); color: #34d399;' : ''}">${deductionLabel}</span>
                    </div>
                    <div class="mitigation-desc">
                        <strong>${t.deficiencyLbl}:</strong> ${origAns === 'a' ? q.options.a[currentLang] : q.options.b[currentLang]}<br>
                        <strong>${t.actionLbl}:</strong> ${getMitigationAction(platName, q.id, q, currentLang)}
                    </div>
                </div>
            `;
            listContainer.appendChild(card);
        }
    });

    if (!hasFindings) {
        const emptyDiv = document.createElement('div');
        emptyDiv.className = 'no-findings-msg';
        emptyDiv.innerHTML = currentLang === 'es' 
            ? '🎉 No se detectaron deficiencias estructurales ni factores de exposición mayor en esta instalación.' 
            : '🎉 No structural deficiencies or major exposure factors were detected in this facility.';
        listContainer.appendChild(emptyDiv);
    }
}

// Activar/desactivar simulación de una mitigación
function toggleMitigation(platName, qId) {
    if (!userMitigations[platName]) {
        userMitigations[platName] = {};
    }
    
    userMitigations[platName][qId] = !userMitigations[platName][qId];
    
    renderPlatformQuestions(platName);
    updatePlatformSimulation();
}

// ==========================================
// COMPONENTES GRÁFICOS Y ANALÍTICA (SVG NATIVO)
// ==========================================

// Variable global para controlar la pestaña activa de analíticas
let currentAnalyticsTab = 'bubble';

// Funciones para gestión del Tooltip flotante global con control de límites de pantalla
function showTooltip(evt, htmlContent) {
    const tooltip = document.getElementById('chart-tooltip');
    if (!tooltip) return;
    tooltip.innerHTML = htmlContent;
    tooltip.style.opacity = '1';
    adjustTooltipPosition(evt, tooltip);
}

function moveTooltip(evt) {
    const tooltip = document.getElementById('chart-tooltip');
    if (!tooltip) return;
    adjustTooltipPosition(evt, tooltip);
}

function adjustTooltipPosition(evt, tooltip) {
    const tooltipWidth = tooltip.offsetWidth || 220;
    const tooltipHeight = tooltip.offsetHeight || 100;
    
    const xOffset = 12;
    const yOffset = 12;
    
    let x = evt.clientX + xOffset;
    let y = evt.clientY + yOffset;
    
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;
    
    // Si se sale por la derecha, lo mostramos a la izquierda del cursor
    if (x + tooltipWidth > viewportWidth - 20) {
        x = evt.clientX - tooltipWidth - xOffset;
    }
    
    // Si se sale por abajo, lo mostramos arriba del cursor
    if (y + tooltipHeight > viewportHeight - 20) {
        y = evt.clientY - tooltipHeight - yOffset;
    }
    
    // Asegurar límites mínimos en pantalla
    if (x < 10) x = 10;
    if (y < 10) y = 10;
    
    tooltip.style.left = `${x}px`;
    tooltip.style.top = `${y}px`;
}

function hideTooltip() {
    const tooltip = document.getElementById('chart-tooltip');
    if (!tooltip) return;
    tooltip.style.opacity = '0';
}

// Limpiar HTML para que sea seguro en los atributos onmouseover de SVG
function cleanTooltipHtml(html) {
    if (!html) return '';
    return html.replace(/\n/g, ' ').replace(/"/g, "'").replace(/'/g, "\\'").trim();
}


// Controlador de pestañas de analítica en el Dashboard
function switchAnalyticsTab(tabType) {
    currentAnalyticsTab = tabType;
    
    // Quitar clases activas
    document.querySelectorAll('.analytics-tab-btn').forEach(btn => btn.classList.remove('active'));
    
    // Activar botón seleccionado
    const btn = document.getElementById(`btn-tab-${tabType}`);
    if (btn) btn.classList.add('active');
    
    // Ocultar todos los contenedores
    const c1 = document.getElementById('bubble-chart-box');
    const c2 = document.getElementById('pareto-chart-box');
    const c3 = document.getElementById('sankey-chart-box');
    if (c1) c1.style.display = 'none';
    if (c2) c2.style.display = 'none';
    if (c3) c3.style.display = 'none';
    
    // Mostrar y renderizar activo
    const targetBox = document.getElementById(`${tabType}-chart-box`);
    if (targetBox) {
        targetBox.style.display = 'flex';
        renderAnalyticsChart(tabType);
    }
}

// Renderizar el gráfico seleccionado
function renderAnalyticsChart(tabType) {
    if (tabType === 'bubble') {
        drawBubbleChart('bubble-chart-box');
    } else if (tabType === 'pareto') {
        drawParetoChart('pareto-chart-box');
    } else if (tabType === 'sankey') {
        drawSankeyDiagram('sankey-chart-box');
    }
}

// 1. DIBUJAR GRÁFICO DE RADAR / ARAÑA (PERFIL DE SALUD)
function drawRadarChart(containerId, values) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';
    
    const width = 540;
    const height = 340;
    const centerX = width / 2;
    const centerY = height / 2;
    const rMax = 110;
    const numAxes = 5;
    
    const categories = currentLang === 'es' ? [
        'Luz Libre / Air Gap\n(Módulo C1)',
        'Redundancia Jacket\n(Módulo C2)',
        'Condición Estructural\n(Módulo C3)',
        'Consecuencias / Fluidos\n(Módulo B1-B2)',
        'Iniciadores / Gatillos\n(Módulo D1-D5)'
    ] : [
        'Air Gap / Clearance\n(Module C1)',
        'Jacket Redundancy\n(Module C2)',
        'Structural Damage\n(Module C3)',
        'Consequence / Fluids\n(Module B1-B2)',
        'Reeval. Initiators\n(Module D1-D5)'
    ];
    
    let svg = `<svg width="100%" height="100%" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg">`;
    
    // Filtro de brillo para la poligonal del radar
    svg += `
    <defs>
        <filter id="radar-glow" x="-20%" y="-20%" width="140%" height="140%">
            <feGaussianBlur stdDeviation="3" result="blur" />
            <feComposite in="SourceGraphic" in2="blur" operator="over" />
        </filter>
    </defs>`;
    
    // Pentágonos concéntricos de fondo (niveles de salud 1 al 5)
    for (let level = 1; level <= 5; level++) {
        const r = rMax * (level / 5);
        let points = [];
        for (let i = 0; i < numAxes; i++) {
            const angle = -Math.PI / 2 + (i * 2 * Math.PI / numAxes);
            const x = centerX + r * Math.cos(angle);
            const y = centerY + r * Math.sin(angle);
            points.push(`${x},${y}`);
        }
        svg += `<polygon points="${points.join(' ')}" fill="none" stroke="rgba(1,39,67,0.08)" stroke-width="1" />`;
        svg += `<circle cx="${centerX}" cy="${centerY}" r="${r}" fill="none" stroke="rgba(1,39,67,0.04)" stroke-dasharray="2 3" />`;
        
        // Números guía
        const lblY = centerY - r;
        svg += `<text x="${centerX + 4}" y="${lblY + 8}" fill="rgba(1,39,67,0.3)" font-size="9" font-weight="600">${level}</text>`;
    }
    
    // Líneas radiales y etiquetas de categorías
    for (let i = 0; i < numAxes; i++) {
        const angle = -Math.PI / 2 + (i * 2 * Math.PI / numAxes);
        const xOuter = centerX + rMax * Math.cos(angle);
        const yOuter = centerY + rMax * Math.sin(angle);
        
        // Ejes
        svg += `<line x1="${centerX}" y1="${centerY}" x2="${xOuter}" y2="${yOuter}" stroke="rgba(1,39,67,0.12)" stroke-width="1" />`;
        
        // Colocación de etiquetas
        const labelDist = 135;
        const xText = centerX + labelDist * Math.cos(angle);
        const yText = centerY + labelDist * Math.sin(angle);
        let textAnchor = 'middle';
        
        if (Math.cos(angle) > 0.1) textAnchor = 'start';
        else if (Math.cos(angle) < -0.1) textAnchor = 'end';
        
        const lines = categories[i].split('\n');
        if (lines.length > 1) {
            svg += `<text x="${xText}" y="${yText - 4}" text-anchor="${textAnchor}" fill="var(--text-main)" font-size="12" font-weight="600">
                <tspan x="${xText}" dy="0">${lines[0]}</tspan>
                <tspan x="${xText}" dy="13" fill="var(--text-muted)" font-size="10.5" font-weight="500">${lines[1]}</tspan>
            </text>`;
        } else {
            svg += `<text x="${xText}" y="${yText + 3}" text-anchor="${textAnchor}" fill="var(--text-main)" font-size="12" font-weight="600">${categories[i]}</text>`;
        }
    }
    
    // Polígono de valores de salud
    let valPoints = [];
    for (let i = 0; i < numAxes; i++) {
        const val = values[i] || 1;
        const angle = -Math.PI / 2 + (i * 2 * Math.PI / numAxes);
        const r = rMax * (val / 5);
        const x = centerX + r * Math.cos(angle);
        const y = centerY + r * Math.sin(angle);
        valPoints.push(`${x},${y}`);
    }
    
    // Dibujar el polígono sombreado
    svg += `<polygon points="${valPoints.join(' ')}" fill="rgba(59, 130, 246, 0.2)" stroke="#3b82f6" stroke-width="2.5" filter="url(#radar-glow)" />`;
    
    // Nodos de los valores para interactividad
    for (let i = 0; i < numAxes; i++) {
        const val = values[i] || 1;
        const angle = -Math.PI / 2 + (i * 2 * Math.PI / numAxes);
        const r = rMax * (val / 5);
        const x = centerX + r * Math.cos(angle);
        const y = centerY + r * Math.sin(angle);
        
        const labelClean = categories[i].replace('\n', ' ');
        const tooltipHtml = `
            <div class="tooltip-title">${labelClean}</div>
            <div class="tooltip-row">
                <span class="tooltip-label">${currentLang === 'es' ? 'Nivel Salud' : 'Health Level'}:</span>
                <span class="tooltip-value" style="color:#60a5fa">${val.toFixed(1)} / 5.0</span>
            </div>
        `;
        
        svg += `<circle cx="${x}" cy="${y}" r="4.5" fill="#3b82f6" stroke="#fff" stroke-width="2" class="svg-interactive-el"
            onmouseover="showTooltip(event, '${cleanTooltipHtml(tooltipHtml)}')"
            onmousemove="moveTooltip(event)"
            onmouseout="hideTooltip()" />`;
    }
    
    svg += `</svg>`;
    container.innerHTML = svg;
}

// 2. DIBUJAR GRÁFICO DE CASCADA (EXPLICACIÓN DEL RIESGO ACUMULADO)
function drawWaterfallChart(containerId, waterfallData, totalPoints) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';
    
    const w = 450;
    const h = 260;
    const padL = 40;
    const padR = 20;
    const padT = 30;
    const padB = 40;
    
    const chartW = w - padL - padR;
    const chartH = h - padT - padB;
    
    // Construir los pasos acumulativos de la cascada
    const steps = [];
    steps.push({ label: 'Base', title: currentLang === 'es' ? 'Riesgo Inicial (Base)' : 'Initial Risk (Base)', value: 0, type: 'base' });
    
    waterfallData.forEach(item => {
        steps.push({ label: item.label, title: item.title, value: item.value, type: 'increment' });
    });
    
    steps.push({ label: 'Total', title: currentLang === 'es' ? 'Puntaje de Probabilidad Final' : 'Final Probability Score', value: totalPoints, type: 'total' });
    
    const numSteps = steps.length;
    const colW = Math.min(32, chartW / (numSteps * 1.5));
    const stepGap = (chartW - (numSteps * colW)) / (numSteps - 1 || 1);
    
    // Encontrar escala máxima del eje Y
    const maxVal = Math.max(totalPoints, 6) + 2;
    const getY = (val) => padT + chartH - (val / maxVal * chartH);
    
    let svg = `<svg width="100%" height="100%" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg">`;
    
    // Líneas de cuadrícula horizontal
    const numGridLines = 5;
    for (let i = 0; i <= numGridLines; i++) {
        const val = Math.round(maxVal * (i / numGridLines));
        const y = getY(val);
        svg += `<line x1="${padL}" y1="${y}" x2="${w - padR}" y2="${y}" stroke="rgba(1,39,67,0.06)" stroke-dasharray="2 3" />`;
        svg += `<text x="${padL - 8}" y="${y + 3}" fill="rgba(1,39,67,0.35)" font-size="9" text-anchor="end">${val}</text>`;
    }
    
    let runningSum = 0;
    
    for (let i = 0; i < numSteps; i++) {
        const step = steps[i];
        const colX = padL + (i * (colW + stepGap));
        
        let yStart = 0;
        let yEnd = 0;
        let fill = '';
        let stroke = '';
        let tooltipContent = '';
        
        if (step.type === 'base') {
            yStart = getY(0);
            yEnd = getY(0);
            fill = 'rgba(148, 163, 184, 0.2)';
            stroke = '#94a3b8';
            tooltipContent = `<div class="tooltip-title">${step.title}</div><div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Puntos' : 'Points'}:</span><span class="tooltip-value">0 pts</span></div>`;
        } else if (step.type === 'increment') {
            const startVal = runningSum;
            runningSum += step.value;
            const endVal = runningSum;
            
            yStart = getY(startVal);
            yEnd = getY(endVal);
            fill = step.value >= 3 ? 'rgba(239, 68, 68, 0.45)' : 'rgba(249, 115, 22, 0.45)';
            stroke = step.value >= 3 ? '#ef4444' : '#f97316';
            
            tooltipContent = `
                <div class="tooltip-title">${step.label} - ${step.title}</div>
                <div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Penalización' : 'Penalty'}:</span><span class="tooltip-value" style="color:${stroke}">+${step.value} pts</span></div>
                <div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Subtotal' : 'Subtotal'}:</span><span class="tooltip-value">${endVal} pts</span></div>`;
        } else if (step.type === 'total') {
            yStart = getY(0);
            yEnd = getY(step.value);
            fill = step.value > 0 ? 'rgba(59, 130, 246, 0.4)' : 'rgba(16, 185, 129, 0.4)';
            stroke = step.value > 0 ? '#60a5fa' : '#10b981';
            
            tooltipContent = `
                <div class="tooltip-title">${step.title}</div>
                <div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Total Probabilidad' : 'Total Probability'}:</span><span class="tooltip-value" style="color:${stroke}">${step.value} pts</span></div>`;
        }
        
        const rectH = Math.abs(yStart - yEnd);
        const rectY = Math.min(yStart, yEnd);
        
        // Línea conectora horizontal punteada
        if (i < numSteps - 1) {
            const nextX = padL + ((i + 1) * (colW + stepGap));
            const yConn = step.type === 'total' ? yStart : yEnd;
            svg += `<line x1="${colX + colW}" y1="${yConn}" x2="${nextX}" y2="${yConn}" stroke="rgba(1,39,67,0.2)" stroke-dasharray="3 3" />`;
        }
        
        // Rectángulo del bloque de cascada
        svg += `<rect x="${colX}" y="${rectY}" width="${colW}" height="${Math.max(2, rectH)}" fill="${fill}" stroke="${stroke}" stroke-width="1.5" rx="3" class="svg-interactive-el"
            onmouseover="showTooltip(event, '${cleanTooltipHtml(tooltipContent)}')"
            onmousemove="moveTooltip(event)"
            onmouseout="hideTooltip()" />`;
            
        // Etiqueta del paso debajo del eje X
        svg += `<text x="${colX + colW / 2}" y="${padT + chartH + 16}" fill="rgba(1,39,67,0.5)" font-size="9" font-weight="500" text-anchor="middle">${step.label}</text>`;
        
        // Etiqueta del valor numérico
        const valText = step.type === 'increment' ? `+${step.value}` : `${step.value}`;
        svg += `<text x="${colX + colW / 2}" y="${rectY - 6}" fill="${stroke}" font-size="9" font-weight="600" text-anchor="middle">${valText}</text>`;
    }
    
    // Eje base X
    svg += `<line x1="${padL}" y1="${padT + chartH}" x2="${w - padR}" y2="${padT + chartH}" stroke="rgba(1,39,67,0.15)" stroke-width="1" />`;
    
    svg += `</svg>`;
    container.innerHTML = svg;
}

// 3. DIBUJAR BUBBLE CHART (DECISIÓN ESTRATÉGICA DE FLOTA)
function drawBubbleChart(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';
    
    const w = 550;
    const h = 300;
    const padL = 65;
    const padR = 50;
    const padT = 40;
    const padB = 45;
    
    const chartW = w - padL - padR;
    const chartH = h - padT - padB;
    
    // Recopilar puntos de datos
    const points = [];
    Object.keys(platformsData).forEach(platName => {
        const metrics = calculatePlatformMetrics(platName, platformsData[platName]);
        const meta = getPlatformMeta(platName);
        
        let numericLof = 2; // Media
        if (metrics.probLevel === 'ALTA') numericLof = 3;
        else if (metrics.probLevel === 'BAJA') numericLof = 1;
        
        points.push({
            name: platName,
            age: meta.age,
            occupancy: meta.occupancy,
            occupancyVal: meta.occupancyVal,
            bpd: meta.bpd,
            lof: numericLof,
            lofPoints: metrics.probLevel,
            cof: metrics.finalCons,
            riskCode: metrics.riskCode,
            riskName: metrics.riskName,
            riskColor: metrics.riskColor,
            coordinate: metrics.coordinate
        });
    });
    
    let svg = `<svg width="100%" height="100%" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg">`;
    
    // Escala del Eje X: Años (0 a 40)
    const maxX = 40;
    const getX = (val) => padL + (val / maxX * chartW);
    
    // Escala del Eje Y: LoF (1 = BAJA, 2 = MEDIA, 3 = ALTA)
    const maxY = 3.5;
    const minY = 0.5;
    const getY = (val) => padT + chartH - ((val - minY) / (maxY - minY) * chartH);
    
    // Cuadrícula del Eje Y (Niveles LoF Cualitativos)
    const lofLabels = currentLang === 'es' 
        ? { 3: 'Alta', 2: 'Media', 1: 'Baja' } 
        : { 3: 'High', 2: 'Medium', 1: 'Low' };
        
    for (let level = 1; level <= 3; level++) {
        const y = getY(level);
        svg += `<line x1="${padL}" y1="${y}" x2="${w - padR}" y2="${y}" stroke="rgba(1,39,67,0.06)" />`;
        svg += `<text x="${padL - 10}" y="${y + 4}" fill="rgba(1,39,67,0.35)" font-size="9" font-weight="500" text-anchor="end">${lofLabels[level]}</text>`;
    }
    
    // Cuadrícula del Eje X (Años)
    for (let age = 0; age <= maxX; age += 5) {
        const x = getX(age);
        svg += `<line x1="${x}" y1="${padT}" x2="${x}" y2="${padT + chartH}" stroke="rgba(1,39,67,0.04)" stroke-dasharray="2 2" />`;
        svg += `<text x="${x}" y="${padT + chartH + 16}" fill="rgba(1,39,67,0.35)" font-size="9" text-anchor="middle">${age} ${currentLang === 'es' ? 'años' : 'yrs'}</text>`;
    }
    
    // Líneas de los ejes X e Y
    svg += `<line x1="${padL}" y1="${padT + chartH}" x2="${w - padR}" y2="${padT + chartH}" stroke="rgba(1,39,67,0.15)" stroke-width="1" />`;
    svg += `<line x1="${padL}" y1="${padT}" x2="${padL}" y2="${padT + chartH}" stroke="rgba(1,39,67,0.15)" stroke-width="1" />`;
    
    // Títulos de ejes
    svg += `<text x="${padL + chartW / 2}" y="${padT + chartH + 34}" fill="rgba(1,39,67,0.45)" font-size="10" font-weight="600" text-anchor="middle">${currentLang === 'es' ? 'Antigüedad (Años en Servicio)' : 'Age (Years in Service)'}</text>`;
    svg += `<text x="14" y="${padT + chartH / 2}" fill="rgba(1,39,67,0.45)" font-size="10" font-weight="600" text-anchor="middle" transform="rotate(-90 14 ${padT + chartH / 2})">${currentLang === 'es' ? 'Probabilidad de Falla (LoF)' : 'Probability of Failure (LoF)'}</text>`;
    
    // Precalcula las coordenadas y radios originales
    points.forEach(p => {
        p.cx_orig = getX(p.age);
        p.cy_orig = getY(p.lof);
        
        let r = 8;
        if (p.occupancy === 'L-1') r = 18;
        else if (p.occupancy === 'L-2') r = 13;
        p.r = r;
        
        p.cx = p.cx_orig;
        p.cy = p.cy_orig;
    });

    // Agrupa las plataformas que se superponen (mismo nivel de LoF y distancia X menor al diámetro de colisión)
    const groups = [];
    points.forEach(p => {
        let foundGroup = null;
        for (let g of groups) {
            let overlap = false;
            for (let member of g) {
                if (member.lof === p.lof) {
                    const dist = Math.abs(member.cx_orig - p.cx_orig);
                    const minDist = member.r + p.r + 10; // Umbral de superposición en píxeles
                    if (dist < minDist) {
                        overlap = true;
                        break;
                    }
                }
            }
            if (overlap) {
                foundGroup = g;
                break;
            }
        }
        if (foundGroup) {
            foundGroup.push(p);
        } else {
            groups.push([p]);
        }
    });

    // Desplaza horizontalmente los círculos dentro de cada grupo para evitar superposición
    groups.forEach(g => {
        if (g.length > 1) {
            // Ordenar por edad y luego por nombre para consistencia
            g.sort((a, b) => {
                if (a.age !== b.age) return a.age - b.age;
                return a.name.localeCompare(b.name);
            });
            
            const gap = 6; // Distancia entre burbujas
            let totalW = 0;
            g.forEach(p => {
                totalW += p.r * 2;
            });
            totalW += gap * (g.length - 1);
            
            // Centrar el grupo sobre la coordenada X promedio original
            const avgCx = g.reduce((sum, p) => sum + p.cx_orig, 0) / g.length;
            
            let currentX = avgCx - totalW / 2;
            g.forEach(p => {
                p.cx = currentX + p.r;
                currentX += p.r * 2 + gap;
            });
        }
    });

    // Dibujar burbujas de plataformas
    points.forEach(p => {
        const cx = p.cx;
        const cy = p.cy;
        const r = p.r;
        
        const strokeColor = p.name === 'Balam-TC' ? '#64748b' : p.riskColor;
        const fillColor = p.name === 'Balam-TC' ? 'rgba(203, 213, 225, 0.6)' : p.riskColor + 'cc';
        const strokeDash = p.name === 'Balam-TC' ? 'stroke-dasharray="3 3"' : '';
        
        const t = uiTranslations[currentLang];
        const tooltipContent = `
            <div class="tooltip-title">${p.name} ${p.name === 'Balam-TC' ? (currentLang === 'es' ? '(Inactiva)' : '(Inactive)') : ''}</div>
            <div class="tooltip-row"><span class="tooltip-label">${t.tooltipAge}:</span><span class="tooltip-value">${p.age} ${currentLang === 'es' ? 'años' : 'years'}</span></div>
            <div class="tooltip-row"><span class="tooltip-label">${t.tooltipLoF}:</span><span class="tooltip-value">${translateLevel(p.lofPoints, 'title')}</span></div>
            <div class="tooltip-row"><span class="tooltip-label">${t.tooltipCoF}:</span><span class="tooltip-value">${p.cof}</span></div>
            <div class="tooltip-row"><span class="tooltip-label">${t.tooltipRisk}:</span><span class="tooltip-value" style="color:${strokeColor}">${p.name === 'Balam-TC' ? (currentLang === 'es' ? 'Inactiva' : 'Inactive') : p.riskName}</span></div>
            <div class="tooltip-row"><span class="tooltip-label">${t.tooltipStaff}:</span><span class="tooltip-value">${p.occupancyVal} pax (${p.occupancy})</span></div>
            <div class="tooltip-row"><span class="tooltip-label">${t.tooltipBpd}:</span><span class="tooltip-value">${p.bpd.toLocaleString()} bpd</span></div>
            <div style="font-size:10px; color:#60a5fa; margin-top:6px; font-weight:600; text-align:center">${currentLang === 'es' ? 'Haz clic para ver plataforma' : 'Click to view platform'}</div>
        `;
        
        // Encontrar grupo y su índice para alternar etiquetas
        const group = groups.find(g => g.includes(p));
        const indexInGroup = group ? group.indexOf(p) : 0;
        
        let textX = cx;
        let textY = cy - r - 5;
        let textAnchor = 'middle';
        
        // Si hay superposición en el grupo, alternamos las etiquetas arriba y abajo
        if (group && group.length > 1) {
            if (indexInGroup % 2 === 1) {
                // Etiqueta abajo
                textY = cy + r + 12;
            }
        }
        
        const labelText = p.name === 'Balam-TC' ? p.name + (currentLang === 'es' ? ' (Inactiva)' : ' (Inactive)') : p.name;

        svg += `
        <g class="svg-interactive-el svg-bubble" style="color: ${strokeColor}; cursor: pointer;" onclick="switchTab('platform:${p.name}')"
            onmouseover="showTooltip(event, '${cleanTooltipHtml(tooltipContent)}')"
            onmousemove="moveTooltip(event)"
            onmouseout="hideTooltip()">
            <circle cx="${cx}" cy="${cy}" r="${r}" fill="${fillColor}" stroke="${strokeColor}" stroke-width="1.5" ${strokeDash} />
            <text x="${textX}" y="${textY}" fill="#0f172a" stroke="#ffffff" stroke-width="2.5" paint-order="stroke fill" font-size="8.5" font-weight="600" text-anchor="${textAnchor}" pointer-events="none">${labelText}</text>
        </g>
        `;
    });
    
    svg += `</svg>`;
    
    let noteHtml = '';
    if (currentLang === 'es') {
        noteHtml = `<div class="chart-note" style="margin-top: 15px; font-size: 0.8rem; color: var(--text-muted); background: var(--surface-container-low); padding: 12px 16px; border-radius: 8px; line-height: 1.5; outline: 1px solid var(--outline-variant); display: flex; gap: 10px; align-items: flex-start;">
            <span style="font-size: 1.1rem; line-height: 1;">⚠️</span>
            <div>
                <strong>Nota Operativa:</strong> La plataforma <strong>Balam-TC</strong> se encuentra fuera de operación y está clasificada como una estructura inactiva (producción de 0 BPD y 0 personal a bordo).
            </div>
        </div>`;
    } else {
        noteHtml = `<div class="chart-note" style="margin-top: 15px; font-size: 0.8rem; color: var(--text-muted); background: var(--surface-container-low); padding: 12px 16px; border-radius: 8px; line-height: 1.5; outline: 1px solid var(--outline-variant); display: flex; gap: 10px; align-items: flex-start;">
            <span style="font-size: 1.1rem; line-height: 1;">⚠️</span>
            <div>
                <strong>Operational Note:</strong> The <strong>Balam-TC</strong> platform is out of service and classified as an inactive structure (0 BPD production and 0 personnel on board).
            </div>
        </div>`;
    }
    
    container.innerHTML = svg + noteHtml;
}

// 4. DIBUJAR DIAGRAMA DE PARETO DE DESVIACIONES
function drawParetoChart(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';
    
    const w = 550;
    const h = 300;
    const padL = 40;
    const padR = 40;
    const padT = 30;
    const padB = 40;
    
    const chartW = w - padL - padR;
    const chartH = h - padT - padB;
    
    // Contar totalizadores de penalización por pregunta
    const questionPenalties = {};
    questionsMetadata.forEach(q => {
        questionPenalties[q.id] = { id: q.id, label: q.label[currentLang], title: q.text[currentLang], points: 0, occurrences: 0 };
    });
    
    Object.keys(platformsData).forEach(platName => {
        const answers = platformsData[platName];
        Object.keys(answers).forEach(qId => {
            if (questionPenalties[qId]) {
                const opt = answers[qId];
                let isFinding = false;
                let pts = 0;
                if (qId.startsWith('D')) {
                    isFinding = (opt === 'a');
                    pts = isFinding ? 2 : 0;
                } else {
                    isFinding = (opt === 'a' || opt === 'b');
                    pts = opt === 'a' ? 3 : (opt === 'b' ? 1 : 0);
                }
                
                if (isFinding) {
                    questionPenalties[qId].points += pts;
                    questionPenalties[qId].occurrences++;
                }
            }
        });
    });
    
    const sortedData = Object.values(questionPenalties).sort((a, b) => b.points - a.points);
    const totalPointsSum = sortedData.reduce((sum, item) => sum + item.points, 0) || 1;
    
    let runningSum = 0;
    sortedData.forEach(item => {
        runningSum += item.points;
        item.cumPercent = (runningSum / totalPointsSum) * 100;
    });
    
    const numBars = sortedData.length;
    const colW = Math.min(24, chartW / (numBars * 1.5));
    const stepGap = (chartW - (numBars * colW)) / (numBars - 1 || 1);
    
    const maxPoints = Math.max(...sortedData.map(item => item.points), 5) + 1;
    
    const getYLeft = (val) => padT + chartH - (val / maxPoints * chartH);
    const getYRight = (val) => padT + chartH - (val / 100 * chartH);
    const getX = (i) => padL + (i * (colW + stepGap));
    
    let svg = `<svg width="100%" height="100%" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg">`;
    
    // Eje Izquierdo: Guías
    const numGridLines = 5;
    for (let i = 0; i <= numGridLines; i++) {
        const pts = Math.round(maxPoints * (i / numGridLines));
        const y = getYLeft(pts);
        svg += `<line x1="${padL}" y1="${y}" x2="${w - padR}" y2="${y}" stroke="rgba(1,39,67,0.06)" stroke-dasharray="2 2" />`;
        svg += `<text x="${padL - 8}" y="${y + 3}" fill="rgba(1,39,67,0.35)" font-size="9" text-anchor="end">${pts}</text>`;
    }
    
    // Eje Derecho: Guías de porcentaje
    for (let i = 0; i <= numGridLines; i++) {
        const pct = Math.round(100 * (i / numGridLines));
        const y = getYRight(pct);
        svg += `<text x="${w - padR + 8}" y="${y + 3}" fill="rgba(96, 165, 250, 0.4)" font-size="9">${pct}%</text>`;
    }
    
    // Pintar barras
    for (let i = 0; i < numBars; i++) {
        const item = sortedData[i];
        const barX = getX(i);
        const barY = getYLeft(item.points);
        const barH = padT + chartH - barY;
        
        const tooltipContent = `
            <div class="tooltip-title">${item.id} - ${item.label}</div>
            <div style="font-size:11px; margin-bottom:6px; color:#cbd5e1">${item.title}</div>
            <div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Puntos Totales:' : 'Total Points:'}</span><span class="tooltip-value" style="color:#ef4444">${item.points} pts</span></div>
            <div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Instalaciones Afectadas:' : 'Affected Facilities:'}</span><span class="tooltip-value">${item.occurrences}</span></div>
            <div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Impacto Acumulado:' : 'Cumulative Impact:'}</span><span class="tooltip-value" style="color:#60a5fa">${item.cumPercent.toFixed(1)}%</span></div>
        `;
        
        svg += `<rect x="${barX}" y="${barY}" width="${colW}" height="${Math.max(1, barH)}" fill="rgba(239, 68, 68, 0.3)" stroke="#ef4444" stroke-width="1.2" rx="2" class="svg-interactive-el"
            onmouseover="showTooltip(event, '${cleanTooltipHtml(tooltipContent)}')"
            onmousemove="moveTooltip(event)"
            onmouseout="hideTooltip()" />`;
            
        // Etiqueta del código abajo
        svg += `<text x="${barX + colW / 2}" y="${padT + chartH + 14}" fill="rgba(1,39,67,0.45)" font-size="9" text-anchor="middle">${item.id}</text>`;
    }
    
    // Dibujar curva acumulada (Pareto)
    let pathD = '';
    for (let i = 0; i < numBars; i++) {
        const item = sortedData[i];
        const x = getX(i) + colW / 2;
        const y = getYRight(item.cumPercent);
        if (i === 0) pathD += `M ${x} ${y}`;
        else pathD += ` L ${x} ${y}`;
    }
    svg += `<path d="${pathD}" fill="none" stroke="#60a5fa" stroke-width="2" />`;
    
    // Puntos interactivos sobre la curva
    for (let i = 0; i < numBars; i++) {
        const item = sortedData[i];
        const x = getX(i) + colW / 2;
        const y = getYRight(item.cumPercent);
        
        const ptTooltip = `<div class="tooltip-title">${currentLang === 'es' ? 'Curva de Pareto' : 'Pareto Curve'}</div><div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Impacto Acumulado' : 'Cumulative Impact'}:</span><span class="tooltip-value" style="color:#60a5fa">${item.cumPercent.toFixed(1)}%</span></div>`;
        svg += `<circle cx="${x}" cy="${y}" r="3.5" fill="#60a5fa" stroke="#fff" stroke-width="1" class="svg-interactive-el"
            onmouseover="showTooltip(event, '${cleanTooltipHtml(ptTooltip)}')"
            onmousemove="moveTooltip(event)"
            onmouseout="hideTooltip()" />`;
    }
    
    // Ejes
    svg += `<line x1="${padL}" y1="${padT + chartH}" x2="${w - padR}" y2="${padT + chartH}" stroke="rgba(1,39,67,0.15)" stroke-width="1" />`;
    
    svg += `<text x="${padL - 10}" y="${padT - 10}" fill="rgba(1,39,67,0.45)" font-size="8" text-anchor="end">${currentLang === 'es' ? 'Puntos' : 'Points'}</text>`;
    svg += `<text x="${w - padR + 10}" y="${padT - 10}" fill="rgba(96, 165, 250, 0.4)" font-size="8">% ${currentLang === 'es' ? 'Acum.' : 'Cum.'}</text>`;
    
    svg += `</svg>`;
    container.innerHTML = svg;
}

// 5. DIBUJAR DIAGRAMA DE SANKEY (FLUJO DE GESTIÓN Y ESCALAMIENTO)
function drawSankeyDiagram(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';
    
    const w = 550;
    const h = 300;
    
    // Contar plataformas por categoría de riesgo
    const counts = { H: 0, M: 0, L: 0 };
    const platformLists = { H: [], M: [], L: [] };
    
    Object.keys(platformsData).forEach(platName => {
        const metrics = calculatePlatformMetrics(platName, platformsData[platName]);
        counts[metrics.riskCode]++;
        platformLists[metrics.riskCode].push(platName);
    });
    
    const totalFlota = Object.keys(platformsData).length;
    
    // Posiciones de columnas
    const xLeft = 60;
    const xMid = 220;
    const xRight = 380;
    const nodeW = 16;
    
    const totalHeight = 180;
    const gap = 16;
    
    const scale = totalHeight / (totalFlota || 1);
    
    const categoriesOrder = ['H', 'M', 'L'];
    const riskLabels = {
        H: { es: 'Alto', en: 'High', color: '#ef4444' },
        M: { es: 'Medio', en: 'Medium', color: '#f59e0b' },
        L: { es: 'Bajo', en: 'Low', color: '#10b981' }
    };
    
    const actionLabels = {
        H: { es: 'Pushover N3 + NDT Sub. 3-5a', en: 'Pushover N3 + NDT Sub. 3-5y' },
        M: { es: 'GVI 3-5a + Monitoreo N2/3', en: 'GVI 3-5y + N2/3 Monitoring' },
        L: { es: 'GVI Extendida (5-6a)', en: 'Extended GVI (5-6y)' }
    };
    
    let svg = `<svg width="100%" height="100%" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg">`;
    
    // Definir gradientes translúcidos para los flujos
    svg += `
    <defs>
        <linearGradient id="grad-H" x1="0" y1="0" x2="1" y2="0">
            <stop offset="0%" stop-color="#94a3b8" stop-opacity="0.25"/>
            <stop offset="100%" stop-color="#ef4444" stop-opacity="0.25"/>
        </linearGradient>
        <linearGradient id="grad-M" x1="0" y1="0" x2="1" y2="0">
            <stop offset="0%" stop-color="#94a3b8" stop-opacity="0.25"/>
            <stop offset="100%" stop-color="#f59e0b" stop-opacity="0.25"/>
        </linearGradient>
        <linearGradient id="grad-L" x1="0" y1="0" x2="1" y2="0">
            <stop offset="0%" stop-color="#94a3b8" stop-opacity="0.25"/>
            <stop offset="100%" stop-color="#10b981" stop-opacity="0.25"/>
        </linearGradient>
        
        <linearGradient id="grad-act-H" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#ef4444" stop-opacity="0.25"/><stop offset="100%" stop-color="#ef4444" stop-opacity="0.25"/></linearGradient>
        <linearGradient id="grad-act-M" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#f59e0b" stop-opacity="0.25"/><stop offset="100%" stop-color="#f59e0b" stop-opacity="0.25"/></linearGradient>
        <linearGradient id="grad-act-L" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#10b981" stop-opacity="0.25"/><stop offset="100%" stop-color="#10b981" stop-opacity="0.25"/></linearGradient>
    </defs>`;
    
    const leftY = 40;
    const leftH = totalFlota * scale;
    
    let currentMidY = leftY;
    const midNodes = {};
    
    categoriesOrder.forEach(cat => {
        const count = counts[cat];
        const nodeH = count * scale;
        midNodes[cat] = {
            y: currentMidY,
            h: nodeH,
            count: count
        };
        if (count > 0) {
            currentMidY += nodeH + gap;
        }
    });
    
    // Curvas Bézier de flujo entre columna 1 y columna 2
    let runningYLeft = leftY;
    categoriesOrder.forEach(cat => {
        const node = midNodes[cat];
        if (node.count === 0) return;
        
        const yStartTop = runningYLeft;
        runningYLeft += node.h;
        const yStartBottom = runningYLeft;
        
        const yEndTop = node.y;
        const yEndBottom = node.y + node.h;
        
        const d = `
            M ${xLeft + nodeW} ${yStartTop}
            C ${(xLeft + xMid) / 2} ${yStartTop}, ${(xLeft + xMid) / 2} ${yEndTop}, ${xMid} ${yEndTop}
            L ${xMid} ${yEndBottom}
            C ${(xLeft + xMid) / 2} ${yEndBottom}, ${(xLeft + xMid) / 2} ${yStartBottom}, ${xLeft + nodeW} ${yStartBottom}
            Z
        `;
        
        const platListStr = platformLists[cat].join(', ');
        const tooltipText = `
            <div class="tooltip-title">${currentLang === 'es' ? 'Flujo de Riesgo' : 'Risk Flow'}</div>
            <div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Plataformas' : 'Platforms'}:</span><span class="tooltip-value">${node.count} / ${totalFlota}</span></div>
            <div style="font-size:11px; margin-top:4px; color:#cbd5e1"><strong>${platListStr}</strong></div>
        `;
        
        svg += `<path d="${d}" fill="url(#grad-${cat})" class="svg-sankey-link" 
            onmouseover="showTooltip(event, '${cleanTooltipHtml(tooltipText)}')"
            onmousemove="moveTooltip(event)"
            onmouseout="hideTooltip()" />`;
    });
    
    // Curvas Bézier de flujo entre columna 2 y columna 3
    categoriesOrder.forEach(cat => {
        const node = midNodes[cat];
        if (node.count === 0) return;
        
        const yStartTop = node.y;
        const yStartBottom = node.y + node.h;
        
        const yEndTop = node.y;
        const yEndBottom = node.y + node.h;
        
        const d = `
            M ${xMid + nodeW} ${yStartTop}
            C ${(xMid + xRight) / 2} ${yStartTop}, ${(xMid + xRight) / 2} ${yEndTop}, ${xRight} ${yEndTop}
            L ${xRight} ${yEndBottom}
            C ${(xMid + xRight) / 2} ${yEndBottom}, ${(xMid + xRight) / 2} ${yStartBottom}, ${xMid + nodeW} ${yStartBottom}
            Z
        `;
        
        const tooltipText = `
            <div class="tooltip-title">${currentLang === 'es' ? 'Acción de Inspección' : 'Inspection Action'}</div>
            <div class="tooltip-row"><span class="tooltip-label">${currentLang === 'es' ? 'Nivel' : 'Level'}:</span><span class="tooltip-value">${riskLabels[cat][currentLang]}</span></div>
            <div style="font-size:11px; margin-top:4px; color:#cbd5e1"><strong>${actionLabels[cat][currentLang]}</strong></div>
        `;
        
        svg += `<path d="${d}" fill="url(#grad-act-${cat})" class="svg-sankey-link"
            onmouseover="showTooltip(event, '${cleanTooltipHtml(tooltipText)}')"
            onmousemove="moveTooltip(event)"
            onmouseout="hideTooltip()" />`;
    });
    
    // Dibujar bloques de nodo columna 1
    svg += `<rect x="${xLeft}" y="${leftY}" width="${nodeW}" height="${leftH}" fill="#94a3b8" rx="2" class="svg-sankey-node" />`;
    svg += `<text x="${xLeft - 8}" y="${leftY + leftH / 2 + 3}" fill="var(--text-main)" font-size="10" font-weight="700" text-anchor="end">${currentLang === 'es' ? 'Flota' : 'Fleet'} (${totalFlota})</text>`;
    
    // Dibujar bloques columna 2
    categoriesOrder.forEach(cat => {
        const node = midNodes[cat];
        if (node.count === 0) return;
        
        const label = riskLabels[cat];
        svg += `<rect x="${xMid}" y="${node.y}" width="${nodeW}" height="${node.h}" fill="${label.color}" rx="2" class="svg-sankey-node" />`;
        svg += `<text x="${xMid + nodeW + 6}" y="${node.y + node.h / 2 + 3}" fill="var(--text-main)" font-size="9" font-weight="600" text-anchor="start">${label[currentLang]} (${node.count})</text>`;
    });
    
    // Dibujar bloques columna 3
    categoriesOrder.forEach(cat => {
        const node = midNodes[cat];
        if (node.count === 0) return;
        
        const label = riskLabels[cat];
        svg += `<rect x="${xRight}" y="${node.y}" width="${nodeW}" height="${node.h}" fill="${label.color}" rx="2" class="svg-sankey-node" />`;
        svg += `<text x="${xRight + nodeW + 8}" y="${node.y + node.h / 2 + 3}" fill="var(--text-muted)" font-size="9" font-weight="500" text-anchor="start">${actionLabels[cat][currentLang]}</text>`;
    });
    
    svg += `</svg>`;
    container.innerHTML = svg;
}

// ==========================================
// MÓDULO DE EXPORTACIÓN A WORD (HTML COMPATIBLE)
// ==========================================

function exportFleetReport() {
    const isES = currentLang === 'es';
    
    // Contar estados de riesgo
    let countTotal = 0;
    let countHigh = 0;
    let countMedium = 0;
    let countLow = 0;
    
    const platforms = [];
    Object.keys(platformsData).forEach(platName => {
        const answers = platformsData[platName];
        const m = calculatePlatformMetrics(platName, answers);
        platforms.push({ name: platName, answers, metrics: m });
        
        countTotal++;
        if (m.riskCode === 'H') countHigh++;
        else if (m.riskCode === 'M') countMedium++;
        else if (m.riskCode === 'L') countLow++;
    });
    
    // Ordenar de mayor a menor criticidad (igual que en el dashboard)
    const riskSortWeight = { 'H': 3, 'M': 2, 'L': 1 };
    platforms.sort((a, b) => {
        const weightA = riskSortWeight[a.metrics.riskCode] || 0;
        const weightB = riskSortWeight[b.metrics.riskCode] || 0;
        if (weightA !== weightB) return weightB - weightA;
        return b.metrics.probPoints - a.metrics.probPoints;
    });

    let html = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
        <meta charset="utf-8">
        <title>${isES ? 'Reporte de Integridad Estructural' : 'Structural Integrity Report'}</title>
        <!--[if gte mso 9]>
        <xml>
            <w:WordDocument>
                <w:View>Print</w:View>
                <w:Zoom>100</w:Zoom>
            </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
            body {
                font-family: Arial, sans-serif;
                color: #333333;
                line-height: 1.4;
                font-size: 11pt;
            }
            @page {
                size: letter;
                margin: 2.54cm;
            }
            h1 {
                font-size: 18pt;
                color: #012743;
                border-bottom: 2px solid #012743;
                padding-bottom: 4px;
                margin-top: 24pt;
                margin-bottom: 12pt;
                page-break-after: avoid;
            }
            h2 {
                font-size: 14pt;
                color: #1d3d5a;
                margin-top: 16pt;
                margin-bottom: 8pt;
                page-break-after: avoid;
            }
            h3 {
                font-size: 12pt;
                color: #333333;
                margin-top: 12pt;
                margin-bottom: 6pt;
                page-break-after: avoid;
            }
            table {
                border-collapse: collapse;
                width: 100%;
                margin-bottom: 16px;
            }
            th, td {
                border: 1px solid #D3D3D3;
                padding: 6px 10px;
                text-align: left;
                font-size: 9.5pt;
            }
            th {
                background-color: #012743;
                color: #ffffff;
                font-weight: bold;
            }
            .bg-gray {
                background-color: #F2F4F6;
            }
            .text-center {
                text-align: center;
            }
            .bold {
                font-weight: bold;
            }
            .risk-cell-H {
                background-color: rgba(239, 68, 68, 0.15);
                color: #990000;
                font-weight: bold;
            }
            .risk-cell-M {
                background-color: rgba(245, 158, 11, 0.15);
                color: #7F6000;
                font-weight: bold;
            }
            .risk-cell-L {
                background-color: rgba(16, 185, 129, 0.15);
                color: #274E13;
                font-weight: bold;
            }
            .callout {
                border-left: 5px solid #EF4444;
                background-color: #FDF2F2;
                padding: 10px 12px;
                margin-bottom: 16px;
                font-style: italic;
            }
            .callout-info {
                border-left: 5px solid #64748B;
                background-color: #F1F5F9;
                padding: 10px 12px;
                margin-bottom: 16px;
                font-style: italic;
            }
            .page-break {
                page-break-before: always;
            }
            .cover {
                text-align: left;
                padding-top: 50px;
                height: 80%;
            }
            .cover-title {
                font-size: 26pt;
                font-weight: bold;
                color: #012743;
                margin-bottom: 12px;
            }
            .cover-subtitle {
                font-size: 14pt;
                color: #64748b;
                margin-bottom: 40px;
                font-style: italic;
            }
            .cover-line {
                border-bottom: 3px solid #012743;
                margin-bottom: 30px;
            }
            .cover-meta {
                margin-top: 100px;
                width: 100%;
            }
            .bullet-list {
                margin-left: 20px;
                margin-bottom: 10px;
            }
        </style>
    </head>
    <body>
        <!-- PORTADA -->
        <div class="cover">
            <div style="font-size: 10pt; font-weight: bold; color: #012743; text-align: right; margin-bottom: 80px;">
                ${isES ? 'ACTIVO DE PRODUCCIÓN EK-BALAM' : 'EK-BALAM PRODUCTION ASSET'}
            </div>
            <div class="cover-line"></div>
            <div class="cover-title">${isES ? 'REPORTE DE VALORACIÓN DE INTEGRIDAD ESTRUCTURAL' : 'STRUCTURAL INTEGRITY ASSESSMENT REPORT'}</div>
            <div class="cover-subtitle">
                ${isES ? 'Jerarquización de Criticidad y Planes de Mitigación Simulación (Módulo API RP 2SIM)' : 'Criticality Ranking and Mitigation Plans Simulation (API RP 2SIM)'}
            </div>
            
            <table class="cover-meta" style="border: none;">
                <tr style="border: none;"><td style="border: none; width: 180px; font-weight: bold;" class="bg-gray">${isES ? 'Área Contractual:' : 'Contract Area:'}</td><td style="border: none;">${isES ? 'Complejo Ek-Balam (Sonda de Campeche, México)' : 'Ek-Balam Complex (Campeche Sound, Mexico)'}</td></tr>
                <tr style="border: none;"><td style="border: none; font-weight: bold;" class="bg-gray">${isES ? 'Contrato:' : 'Contract:'}</td><td style="border: none;">CNH-M1-EK-BALAM/2017</td></tr>
                <tr style="border: none;"><td style="border: none; font-weight: bold;" class="bg-gray">${isES ? 'Metodología:' : 'Methodology:'}</td><td style="border: none;">${isES ? 'Matriz 3x3 de Criticidad / Norma API RP 2SIM' : '3x3 Criticality Matrix / API RP 2SIM Standard'}</td></tr>
                <tr style="border: none;"><td style="border: none; font-weight: bold;" class="bg-gray">${isES ? 'Fecha de Generación:' : 'Generation Date:'}</td><td style="border: none;">${new Date().toLocaleDateString(isES ? 'es-MX' : 'en-US')} (Simulación en Cliente)</td></tr>
            </table>
        </div>
        
        <div class="page-break"></div>
        
        <!-- RESUMEN EJECUTIVO -->
        <h1>${isES ? '1. Resumen Ejecutivo de la Flota' : '1. Executive Fleet Summary'}</h1>
        <p>
            ${isES 
                ? 'Este documento contiene el reporte de simulación de integridad estructural de las plataformas costa afuera del Complejo de Infraestructura Ek-Balam. Las valoraciones reflejan los parámetros vigentes y las mitigaciones simuladas activamente en la aplicación de escritorio al momento de la exportación.'
                : 'This document contains the structural integrity simulation report for the offshore platforms of the Ek-Balam Infrastructure Complex. The assessments reflect the current parameters and active simulated mitigations in the desktop application at the time of export.'}
        </p>
        
        <h2>${isES ? 'Resumen Estadístico de Riesgo' : 'Risk Statistical Summary'}</h2>
        <table>
            <thead>
                <tr>
                    <th class="text-center">${isES ? 'Métrica de Flota' : 'Fleet Metric'}</th>
                    <th class="text-center">${isES ? 'Riesgo Alto (Nivel 1)' : 'High Risk (Level 1)'}</th>
                    <th class="text-center">${isES ? 'Riesgo Medio (Nivel 2)' : 'Medium Risk (Level 2)'}</th>
                    <th class="text-center">${isES ? 'Riesgo Bajo (Nivel 3)' : 'Low Risk (Level 3)'}</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="bold bg-gray">${isES ? 'Cantidad de Plataformas' : 'Number of Platforms'}</td>
                    <td class="text-center bold" style="background-color: #FDF2F2; color: #EF4444;">${countHigh}</td>
                    <td class="text-center bold" style="background-color: #FFFBEB; color: #F59E0B;">${countMedium}</td>
                    <td class="text-center bold" style="background-color: #ECFDF5; color: #10B981;">${countLow}</td>
                </tr>
            </tbody>
        </table>
        
        <h2>${isES ? 'Jerarquización Priorizada de Activos' : 'Asset Prioritized Hierarchy'}</h2>
        <table>
            <thead>
                <tr>
                    <th>${isES ? 'Plataforma' : 'Platform'}</th>
                    <th class="text-center">${isES ? 'Exposición (CoF)' : 'Exposure (CoF)'}</th>
                    <th class="text-center">${isES ? 'Probabilidad (LoF)' : 'Likelihood (LoF)'}</th>
                    <th class="text-center">${isES ? 'Coordenada Matrix' : 'Matrix Coordinate'}</th>
                    <th class="text-center">${isES ? 'Nivel de Riesgo' : 'Risk Level'}</th>
                    <th class="text-center">${isES ? 'Gatillo N2/N3' : 'Trigger N2/N3'}</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    platforms.forEach(p => {
        const m = p.metrics;
        const triggerText = m.triggerEscalation ? (isES ? 'SÍ (Activo)' : 'YES (Active)') : (isES ? 'No' : 'No');
        const triggerStyle = m.triggerEscalation ? 'style="background-color: #FDE8E8; color: #EF4444; font-weight: bold;"' : '';
        const inactiveText = p.name === 'Balam-TC' ? (isES ? ' (INACTIVA)' : ' (INACTIVE)') : '';
        
        html += `
            <tr>
                <td class="bold">${p.name}${inactiveText}</td>
                <td class="text-center">${m.finalCons}</td>
                <td class="text-center">${translateLevel(m.probLevel, 'title')}</td>
                <td class="text-center">${translateCoordinate(m.coordinate)}</td>
                <td class="text-center risk-cell-${m.riskCode}">${m.riskName}</td>
                <td class="text-center" ${triggerStyle}>${triggerText}</td>
            </tr>
        `;
    });
    
    html += `
            </tbody>
        </table>
        
        <div class="page-break"></div>
        
        <!-- DETALLE POR PLATAFORMA -->
        <h1>${isES ? '2. Diagnóstico Detallado por Plataforma' : '2. Detailed Diagnostic by Platform'}</h1>
    `;
    
    platforms.forEach(p => {
        const platName = p.name;
        const answers = p.answers;
        const m = p.metrics;
        const meta = getPlatformMeta(platName);
        const desc = platformDescriptions[platName] ? platformDescriptions[platName][isES ? 'es' : 'en'] : (isES ? 'Plataforma Satélite de Extracción' : 'Satellite Extraction Platform');
        const intervals = getInspectionIntervals(m.riskCode);
        
        const structureTypeRaw = (answers && answers['TIPO']) ? answers['TIPO'] 
                            : (platName.includes('Hab') ? 'Tetrápodo de Acero'
                            : (platName.includes('TA') && !platName.includes('Balam') ? 'Trípode Adosado'
                            : (platName.includes('TB') ? 'Trípode Adosado'
                            : (platName.includes('TC') ? 'Trípode'
                            : (platName.includes('1') ? (isES ? 'Estructura Marina Ligera' : 'Light Sea Horse Structure') : 'Octápodo de Acero')))));
        const structureType = isES ? structureTypeRaw : translateStructureType(structureTypeRaw);
        const waterDepth = (answers && answers['TIRANTE']) 
                            ? (String(answers['TIRANTE']).endsWith('m') ? answers['TIRANTE'] : `${answers['TIRANTE']}m`)
                            : (platName.includes('TA') ? '51.5m' : (platName.includes('TB') ? '51.5m' : (platName.includes('A') ? '50.0m' : '52.0m')));
        
        html += `
            <div style="margin-bottom: 40px; page-break-inside: avoid;">
                <h2>${platName}</h2>
                <p><span class="bold">${isES ? 'Descripción Operativa:' : 'Operational Description:'}</span> ${desc}</p>
        `;
        
        if (platName === 'Balam-TC') {
            html += `
                <div class="callout-info">
                    ${isES 
                        ? '⚠️ <strong>Atención:</strong> Esta plataforma se encuentra actualmente <strong>fuera de operación / inactiva</strong>. Los parámetros mostrados reflejan su condición inactiva.'
                        : '⚠️ <strong>Attention:</strong> This platform is currently <strong>out of operation / inactive</strong>. The parameters displayed reflect its inactive condition.'}
                </div>
            `;
        }
        
        if (m.triggerEscalation) {
            html += `
                <div class="callout">
                    ${isES 
                        ? '⚠️ <strong>Gatillo de Evaluación Activo:</strong> Se han detectado iniciadores de evaluación estructural activos (Módulo D) en esta instalación. Se requiere escalar la estructura inmediatamente a un análisis de ingeniería detallado de Nivel II/III.' 
                        : '⚠️ <strong>Evaluation Trigger Active:</strong> Active structural evaluation initiators (Module D) have been detected. Immediate escalation to Level II/III detailed analysis is required.'}
                </div>
            `;
        }
        
        html += `
            <h3>${isES ? 'Parámetros Físicos y de Diseño' : 'Physical & Design Parameters'}</h3>
            <table>
                <tr>
                    <td class="bold bg-gray" style="width: 25%;">${isES ? 'Año Instalación:' : 'Installation Year:'}</td>
                    <td style="width: 25%;">${2026 - meta.age}</td>
                    <td class="bold bg-gray" style="width: 25%;">${isES ? 'Antigüedad (2026):' : 'Age (2026):'}</td>
                    <td style="width: 25%;">${meta.age} ${isES ? 'años' : 'years'}</td>
                </tr>
                <tr>
                    <td class="bold bg-gray">${isES ? 'Tirante de Agua:' : 'Water Depth:'}</td>
                    <td>${waterDepth}</td>
                    <td class="bold bg-gray">${isES ? 'Tipo Estructura:' : 'Structure Type:'}</td>
                    <td>${structureType}</td>
                </tr>
                <tr>
                    <td class="bold bg-gray">${isES ? 'Producción Diaria:' : 'Daily Production:'}</td>
                    <td>${meta.bpd.toLocaleString()} BPD</td>
                    <td class="bold bg-gray">${isES ? 'Personal Ordinario:' : 'Personnel on Board:'}</td>
                    <td>${meta.occupancyVal} POB</td>
                </tr>
            </table>
            
            <h3>${isES ? 'Resultados de Evaluación de Riesgos (API RP 2SIM)' : 'Risk Assessment Results (API RP 2SIM)'}</h3>
            <table>
                <thead>
                    <tr>
                        <th class="text-center">${isES ? 'Probabilidad (LoF)' : 'Likelihood (LoF)'}</th>
                        <th class="text-center">${isES ? 'Exposición (CoF)' : 'Exposure (CoF)'}</th>
                        <th class="text-center">${isES ? 'Coordenada Matrix' : 'Matrix Coordinate'}</th>
                        <th class="text-center">${isES ? 'Riesgo Resultante' : 'Resulting Risk'}</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="text-center bold">${translateLevel(m.probLevel, 'title')}</td>
                        <td class="text-center bold">${m.finalCons}</td>
                        <td class="text-center bold">${translateCoordinate(m.coordinate)}</td>
                        <td class="text-center risk-cell-${m.riskCode}">${m.riskName}</td>
                    </tr>
                </tbody>
            </table>
            
            <h3>${isES ? 'Cuestionario y Estatus de Desviaciones' : 'Questionnaire & Deviations Status'}</h3>
            <table>
                <thead>
                    <tr>
                        <th style="width: 10%;" class="text-center">ID</th>
                        <th style="width: 55%;">${isES ? 'Criterio Evaluado' : 'Assessed Criterion'}</th>
                        <th style="width: 35%;">${isES ? 'Respuesta Seleccionada' : 'Selected Answer'}</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        Object.keys(answers).forEach(qId => {
            const val = answers[qId];
            const q = questionsMetadata.find(item => item.id === qId);
            if (!q) return;
            
            let ansText = "N/D";
            let isDef = false;
            if (val === 'a') {
                ansText = q.options.a[currentLang];
                isDef = true;
            } else if (val === 'b') {
                ansText = q.options.b[currentLang];
                if (qId.startsWith('C')) isDef = true;
            } else if (val === 'c') {
                ansText = q.options.c[currentLang];
            }
            
            const cellBg = isDef ? 'style="background-color: #FDF2F2; color: #EF4444; font-weight: bold;"' : 'style="background-color: #ECFDF5; color: #274E13;"';
            
            html += `
                <tr>
                    <td class="text-center bold bg-gray">${qId}</td>
                    <td>${q.text[currentLang]}</td>
                    <td ${cellBg}>${ansText}</td>
                </tr>
            `;
        });
        
        html += `
                </tbody>
            </table>
            
            <h3>${isES ? 'Plan de Mitigación y Acciones Correctivas' : 'Mitigation Plan & Corrective Actions'}</h3>
        `;
        
        let hasMit = false;
        let mitHtml = '<ul class="bullet-list">';
        Object.keys(answers).forEach(qId => {
            const val = answers[qId];
            const q = questionsMetadata.find(item => item.id === qId);
            if (!q) return;
            
            let isDef = false;
            if (val === 'a') isDef = true;
            else if (val === 'b' && qId.startsWith('C')) isDef = true;
            
            if (isDef) {
                hasMit = true;
                mitHtml += `<li><span class="bold" style="color: #EF4444;">[${qId}]</span> ${getMitigationAction(platName, qId, q, currentLang)}</li>`;
            }
        });
        mitHtml += '</ul>';
        
        if (hasMit) {
            html += mitHtml;
        } else {
            html += `<p style="color: #274E13; font-weight: bold;">${isES ? '🏆 La plataforma se encuentra en estado óptimo. No se requieren acciones correctivas.' : '🏆 The platform is in optimal state. No corrective actions are required.'}</p>`;
        }
        
        html += `
            <h3>${isES ? 'Intervalos Máximos de Inspección Recomendados' : 'Maximum Recommended Inspection Intervals'}</h3>
            <table>
                <thead>
                    <tr>
                        <th class="text-center">${isES ? 'Visual General (GVI)' : 'General Visual (GVI)'}</th>
                        <th class="text-center">${isES ? 'Visual Detallada (CVI)' : 'Close Visual (CVI)'}</th>
                        <th class="text-center">${isES ? 'Ensayos No Destructivos (NDT)' : 'Non-Destructive Testing (NDT)'}</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="text-center bold">${intervals.gvi[currentLang]}</td>
                        <td class="text-center bold">${intervals.cvi[currentLang]}</td>
                        <td class="text-center bold">${intervals.ndt[currentLang]}</td>
                    </tr>
                </tbody>
            </table>
            <hr style="border: none; border-top: 1px dashed #D3D3D3; margin-top: 30px;">
        </div>
        `;
    });
    
    html += `
    </body>
    </html>
    `;
    
    // Descargar
    const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Reporte_Flota_Integridad_Estructural_${isES ? 'Completo' : 'Complete'}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Exportar reporte individual de la plataforma activa
function exportPlatformReportActive() {
    const platName = currentTab.substring(9);
    if (!platName) return;
    
    const isES = currentLang === 'es';
    const answers = platformsData[platName];
    const m = calculatePlatformMetrics(platName, answers);
    const meta = getPlatformMeta(platName);
    const desc = platformDescriptions[platName] ? platformDescriptions[platName][isES ? 'es' : 'en'] : (isES ? 'Plataforma Satélite de Extracción' : 'Satellite Extraction Platform');
    const intervals = getInspectionIntervals(m.riskCode);
    
    const structureTypeRaw = (answers && answers['TIPO']) ? answers['TIPO'] 
                        : (platName.includes('Hab') ? 'Tetrápodo de Acero'
                        : (platName.includes('TA') && !platName.includes('Balam') ? 'Trípode Adosado'
                        : (platName.includes('TB') ? 'Trípode Adosado'
                        : (platName.includes('TC') ? 'Trípode'
                        : (platName.includes('1') ? (isES ? 'Estructura Marina Ligera' : 'Light Sea Horse Structure') : 'Octápodo de Acero')))));
    const structureType = isES ? structureTypeRaw : translateStructureType(structureTypeRaw);
    const waterDepth = (answers && answers['TIRANTE']) 
                        ? (String(answers['TIRANTE']).endsWith('m') ? answers['TIRANTE'] : `${answers['TIRANTE']}m`)
                        : (platName.includes('TA') ? '51.5m' : (platName.includes('TB') ? '51.5m' : (platName.includes('A') ? '50.0m' : '52.0m')));
    const yearInst = 2026 - meta.age;

    let html = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
        <meta charset="utf-8">
        <title>${platName} - ${isES ? 'Reporte de Integridad Estructural' : 'Structural Integrity Report'}</title>
        <!--[if gte mso 9]>
        <xml>
            <w:WordDocument>
                <w:View>Print</w:View>
                <w:Zoom>100</w:Zoom>
            </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
            body {
                font-family: Arial, sans-serif;
                color: #333333;
                line-height: 1.4;
                font-size: 11pt;
            }
            @page {
                size: letter;
                margin: 2.54cm;
            }
            h1 {
                font-size: 22pt;
                color: #012743;
                border-bottom: 3px solid #012743;
                padding-bottom: 5px;
                margin-top: 0;
                margin-bottom: 12pt;
            }
            h2 {
                font-size: 14pt;
                color: #1d3d5a;
                margin-top: 20pt;
                margin-bottom: 8pt;
                page-break-after: avoid;
            }
            table {
                border-collapse: collapse;
                width: 100%;
                margin-bottom: 16px;
            }
            th, td {
                border: 1px solid #D3D3D3;
                padding: 6px 10px;
                text-align: left;
                font-size: 9.5pt;
            }
            th {
                background-color: #012743;
                color: #ffffff;
                font-weight: bold;
            }
            .bg-gray {
                background-color: #F2F4F6;
            }
            .bold {
                font-weight: bold;
            }
            .text-center {
                text-align: center;
            }
            .risk-cell-H {
                background-color: rgba(239, 68, 68, 0.15);
                color: #990000;
                font-weight: bold;
            }
            .risk-cell-M {
                background-color: rgba(245, 158, 11, 0.15);
                color: #7F6000;
                font-weight: bold;
            }
            .risk-cell-L {
                background-color: rgba(16, 185, 129, 0.15);
                color: #274E13;
                font-weight: bold;
            }
            .callout {
                border-left: 5px solid #EF4444;
                background-color: #FDF2F2;
                padding: 10px 12px;
                margin-bottom: 16px;
                font-style: italic;
            }
            .callout-info {
                border-left: 5px solid #64748B;
                background-color: #F1F5F9;
                padding: 10px 12px;
                margin-bottom: 16px;
                font-style: italic;
            }
            .bullet-list {
                margin-left: 20px;
                margin-bottom: 10px;
            }
        </style>
    </head>
    <body>
        <div style="font-size: 9pt; font-weight: bold; color: #012743; text-align: right; margin-bottom: 30px;">
            ${isES ? 'REPORTE DE VALORACIÓN INDIVIDUAL - SIMULACIÓN' : 'INDIVIDUAL ASSESSMENT REPORT - SIMULATION'}
        </div>
        
        <h1>${isES ? 'Plataforma ' : 'Platform '}${platName}</h1>
        <p><span class="bold">${isES ? 'Descripción Operativa:' : 'Operational Description:'}</span> ${desc}</p>
        
        <table style="border: none; margin-top: 15px; margin-bottom: 25px;">
            <tr style="border: none;"><td style="border: none; width: 150px; font-weight: bold;" class="bg-gray">${isES ? 'Área Contractual:' : 'Contract Area:'}</td><td style="border: none;">Complejo Ek-Balam (CNH-M1-EK-BALAM/2017)</td></tr>
            <tr style="border: none;"><td style="border: none; font-weight: bold;" class="bg-gray">${isES ? 'Metodología:' : 'Methodology:'}</td><td style="border: none;">API RP 2SIM (Gestión de Integridad Estructural)</td></tr>
            <tr style="border: none;"><td style="border: none; font-weight: bold;" class="bg-gray">${isES ? 'Fecha de Simulación:' : 'Simulation Date:'}</td><td style="border: none;">${new Date().toLocaleDateString(isES ? 'es-MX' : 'en-US')}</td></tr>
        </table>
        
        ${platName === 'Balam-TC' ? `
            <div class="callout-info">
                ${isES 
                    ? '⚠️ <strong>Atención:</strong> Esta plataforma se encuentra actualmente <strong>fuera de operación / inactiva</strong>. Los parámetros mostrados reflejan su condición inactiva.'
                    : '⚠️ <strong>Attention:</strong> This platform is currently <strong>out of operation / inactive</strong>. The parameters displayed reflect its inactive condition.'}
            </div>
        ` : ''}
        
        ${m.triggerEscalation ? `
            <div class="callout">
                ${isES 
                    ? '⚠️ <strong>Gatillo de Evaluación Activo:</strong> Se han detectado iniciadores de evaluación estructural activos (Módulo D) en esta instalación. Se requiere escalar la estructura inmediatamente a un análisis de ingeniería detallado de Nivel II/III.' 
                    : '⚠️ <strong>Evaluation Trigger Active:</strong> Active structural evaluation initiators (Module D) have been detected. Immediate escalation to Level II/III detailed analysis is required.'}
            </div>
        ` : ''}
        
        <h2>${isES ? '1. Datos de Diseño y Operación' : '1. Design & Operational Data'}</h2>
        <table>
            <tr>
                <td class="bold bg-gray" style="width: 25%;">${isES ? 'Año Instalación:' : 'Installation Year:'}</td>
                <td style="width: 25%;">${yearInst}</td>
                <td class="bold bg-gray" style="width: 25%;">${isES ? 'Antigüedad (2026):' : 'Age (2026):'}</td>
                <td style="width: 25%;">${meta.age} ${isES ? 'años' : 'years'}</td>
            </tr>
            <tr>
                <td class="bold bg-gray">${isES ? 'Tirante de Agua:' : 'Water Depth:'}</td>
                <td>${waterDepth}</td>
                <td class="bold bg-gray">${isES ? 'Tipo Estructura:' : 'Structure Type:'}</td>
                <td>${structureType}</td>
            </tr>
            <tr>
                <td class="bold bg-gray">${isES ? 'Producción Diaria:' : 'Daily Production:'}</td>
                <td>${meta.bpd.toLocaleString()} BPD</td>
                <td class="bold bg-gray">${isES ? 'Personal Ordinario:' : 'Personnel on Board:'}</td>
                <td>${meta.occupancyVal} POB</td>
            </tr>
        </table>
        
        <h2>${isES ? '2. Diagnóstico del Nivel de Riesgo' : '2. Risk Level Diagnostic'}</h2>
        <table>
            <thead>
                <tr>
                    <th class="text-center">${isES ? 'Probabilidad (LoF)' : 'Likelihood (LoF)'}</th>
                    <th class="text-center">${isES ? 'Exposición (CoF)' : 'Exposure (CoF)'}</th>
                    <th class="text-center">${isES ? 'Coordenada Matrix' : 'Matrix Coordinate'}</th>
                    <th class="text-center">${isES ? 'Riesgo Resultante' : 'Resulting Risk'}</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="text-center bold">${translateLevel(m.probLevel, 'title')}</td>
                    <td class="text-center bold">${m.finalCons}</td>
                    <td class="text-center bold">${translateCoordinate(m.coordinate)}</td>
                    <td class="text-center risk-cell-${m.riskCode}">${m.riskName}</td>
                </tr>
            </tbody>
        </table>
        
        <h2>${isES ? '3. Cuestionario Evaluado' : '3. Evaluated Questionnaire'}</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 10%;" class="text-center">ID</th>
                    <th style="width: 55%;">${isES ? 'Criterio Evaluado' : 'Assessed Criterion'}</th>
                    <th style="width: 35%;">${isES ? 'Respuesta Seleccionada' : 'Selected Answer'}</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    Object.keys(answers).forEach(qId => {
        const val = answers[qId];
        const q = questionsMetadata.find(item => item.id === qId);
        if (!q) return;
        
        let ansText = "N/D";
        let isDef = false;
        if (val === 'a') {
            ansText = q.options.a[currentLang];
            isDef = true;
        } else if (val === 'b') {
            ansText = q.options.b[currentLang];
            if (qId.startsWith('C')) isDef = true;
        } else if (val === 'c') {
            ansText = q.options.c[currentLang];
        }
        
        const cellBg = isDef ? 'style="background-color: #FDF2F2; color: #EF4444; font-weight: bold;"' : 'style="background-color: #ECFDF5; color: #274E13;"';
        
        html += `
            <tr>
                <td class="text-center bold bg-gray">${qId}</td>
                <td>${q.text[currentLang]}</td>
                <td ${cellBg}>${ansText}</td>
            </tr>
        `;
    });
    
    html += `
            </tbody>
        </table>
        
        <h2>${isES ? '4. Recomendaciones de Mitigación' : '4. Mitigation Recommendations'}</h2>
    `;
    
    let hasMit = false;
    let mitHtml = '<ul class="bullet-list">';
    Object.keys(answers).forEach(qId => {
        const val = answers[qId];
        const q = questionsMetadata.find(item => item.id === qId);
        if (!q) return;
        
        let isDef = false;
        if (val === 'a') isDef = true;
        else if (val === 'b' && qId.startsWith('C')) isDef = true;
        
        if (isDef) {
            hasMit = true;
            mitHtml += `<li><span class="bold" style="color: #EF4444;">[${qId}]</span> ${getMitigationAction(platName, qId, q, currentLang)}</li>`;
        }
    });
    mitHtml += '</ul>';
    
    if (hasMit) {
        html += mitHtml;
    } else {
        html += `<p style="color: #274E13; font-weight: bold;">${isES ? '🏆 La plataforma se encuentra en estado óptimo. No se requieren acciones correctivas.' : '🏆 The platform is in optimal state. No corrective actions are required.'}</p>`;
    }
    
    html += `
        <h2>${isES ? '5. Intervalos de Inspección Recomendados' : '5. Recommended Inspection Intervals'}</h2>
        <table>
            <thead>
                <tr>
                    <th class="text-center">${isES ? 'Visual General (GVI)' : 'General Visual (GVI)'}</th>
                    <th class="text-center">${isES ? 'Visual Detallada (CVI)' : 'Close Visual (CVI)'}</th>
                    <th class="text-center">${isES ? 'Ensayos No Destructivos (NDT)' : 'Non-Destructive Testing (NDT)'}</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td class="text-center bold">${intervals.gvi[currentLang]}</td>
                    <td class="text-center bold">${intervals.cvi[currentLang]}</td>
                    <td class="text-center bold">${intervals.ndt[currentLang]}</td>
                </tr>
            </tbody>
        </table>
    </body>
    </html>
    `;
    
    // Descargar
    const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Reporte_${platName}_Integridad_Estructural.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ══════════════════════════════════════════════════════════════════
// SUITE SEGURIDAD & SESIÓN (PBKDF2 NATIVO & LOCALSTORAGE)
// ══════════════════════════════════════════════════════════════════

// Base de datos de usuarios por defecto (Suite RBI-Q & SIM-Q unificada)
const defaultUsers = {
    "admin": {
        "salt": "427120d2c718e6b70b9343ab773f6046",
        "hash": "251dce48f5b3c95c121500507907d6f1b39452e33ff9f74b970090baacc2b6cb",
        "role": "admin",
        "full_name": "Michele Leccese"
    },
    "cliente_ekbalam1": {
        "salt": "062b7ae8ab473ca5b30252f24eb99370",
        "hash": "e081c21aae8fab2899cc2f9b17603af80f8ea956c4e56d39c891bf931e834e01",
        "role": "client",
        "full_name": "Sierra Madre Resource Usuario 1"
    },
    "cliente_ekbalam2": {
        "salt": "5e15f17a646b9e4798ebf574f946d82d",
        "hash": "1edbbea81d40cce879bffc8c49d4e1da1db6fe33a4c9ec9c2ebc67676c19b6e2",
        "role": "client",
        "full_name": "Sierra Madre Resource Usuario 2"
    },
    "mleccese": {
        "salt": "0aec9122613372b4a85579f8607983d8",
        "hash": "8861fd847f48a6f121059eaab20ed0f0834614f1e8e22862651a3f55b35b7cad",
        "role": "client",
        "full_name": "Michele Leccese"
    },
    "myanez": {
        "salt": "d49c82840019e43606f518335fc29fb8",
        "hash": "cfaa234fd5f49e40a011867dd08c04e01e05239317d9b413339631261181cb2e",
        "role": "client",
        "full_name": "Medardo Yañez"
    },
    "lfernandez": {
        "salt": "ed3d0aad8eb74a6392369241df512db6",
        "hash": "6a75be6604cfd8dddeb09dad03733bd42015ed0e529014383455a3223e7e896d",
        "role": "client",
        "full_name": "Luis Fernandez"
    },
    "afernandez": {
        "salt": "50119fd117d81ca6774ad54aad513d9a",
        "hash": "370274322643e4143d0dda616e58fafc5ea80480279652f24244f2f3be395f13",
        "role": "client",
        "full_name": "Alis Fernandez"
    },
    "jfarinas": {
        "salt": "1a1bad424e0fc3be9de3e640d328d974",
        "hash": "c31a180f669af868deb6d788ad2641564c32646e103164898f964bf187e78310",
        "role": "client",
        "full_name": "Jose Fariñas"
    },
    "gvegas": {
        "salt": "4eebc1005ebd9400542c20f7977efc5b",
        "hash": "4576eb3dee8c42822bba5d10c5e34897555dc94fa1a0dcdffd0629949faeb7d5",
        "role": "client",
        "full_name": "Gustavo Vegas"
    }
};

// Estado de sesión
let authenticated = false;
let activeUser = '';
let activeRole = '';
let activeFullName = '';

// Convertir cadena hexadecimal a bytes (Uint8Array)
function hexToBytes(hex) {
    let bytes = [];
    for (let c = 0; c < hex.length; c += 2) {
        bytes.push(parseInt(hex.substr(c, 2), 16));
    }
    return new Uint8Array(bytes);
}

// Convertir bytes (Uint8Array) a cadena hexadecimal
function bytesToHex(bytes) {
    return Array.from(new Uint8Array(bytes))
        .map(b => b.toString(16).padStart(2, "0"))
        .join("");
}

// Cifrar contraseña mediante PBKDF2-HMAC-SHA256 (Web Crypto API)
async function hashPassword(password) {
    const saltBytes = window.crypto.getRandomValues(new Uint8Array(16));
    const saltHex = bytesToHex(saltBytes);
    const encoder = new TextEncoder();
    const baseKey = await window.crypto.subtle.importKey(
        "raw",
        encoder.encode(password),
        "PBKDF2",
        false,
        ["deriveBits", "deriveKey"]
    );
    const derivedBits = await window.crypto.subtle.deriveBits(
        {
            name: "PBKDF2",
            salt: saltBytes,
            iterations: 100000,
            hash: "SHA-256"
        },
        baseKey,
        256 // 32 bytes = 256 bits
    );
    const hashHex = bytesToHex(derivedBits);
    return { salt: saltHex, hash: hashHex };
}

// Verificar contraseña comparando PBKDF2 hash
async function verifyPassword(password, saltHex, storedHashHex) {
    try {
        const encoder = new TextEncoder();
        const baseKey = await window.crypto.subtle.importKey(
            "raw",
            encoder.encode(password),
            "PBKDF2",
            false,
            ["deriveBits", "deriveKey"]
        );
        const saltBytes = hexToBytes(saltHex);
        const derivedBits = await window.crypto.subtle.deriveBits(
            {
                name: "PBKDF2",
                salt: saltBytes,
                iterations: 100000,
                hash: "SHA-256"
            },
            baseKey,
            256
        );
        const checkHashHex = bytesToHex(derivedBits);
        return checkHashHex === storedHashHex;
    } catch (e) {
        console.error("Error al verificar contraseña mediante Web Crypto:", e);
        return false;
    }
}

// Cargar base de datos de usuarios
function loadUsers() {
    let stored = localStorage.getItem("simq_users");
    if (stored) {
        try {
            const users = JSON.parse(stored);
            // Migrar nombre si tenía el valor por defecto anterior
            if (users.admin && users.admin.full_name === "Miguel Aguero") {
                users.admin.full_name = "Michele Leccese";
                saveUsers(users);
            }
            return users;
        } catch (e) {
            console.error("Error al decodificar usuarios de localStorage:", e);
        }
    }
    return { ...defaultUsers };
}

// Guardar base de datos de usuarios
function saveUsers(users) {
    localStorage.setItem("simq_users", JSON.stringify(users));
}

// Iniciar sesión (Action Handler)
async function handleLogin(e) {
    e.preventDefault();
    const usernameInput = document.getElementById('usernameInput').value.trim().toLowerCase();
    const passwordInput = document.getElementById('passwordInput').value;
    const errorDiv = document.getElementById('loginErrorMsg');
    
    errorDiv.style.display = 'none';
    
    const users = loadUsers();
    const uData = users[usernameInput];
    
    if (!uData) {
        errorDiv.innerText = currentLang === 'es' ? '❌ Usuario o contraseña incorrectos.' : '❌ Incorrect username or password.';
        errorDiv.style.display = 'block';
        return;
    }
    
    const isValid = await verifyPassword(passwordInput, uData.salt, uData.hash);
    
    if (isValid) {
        authenticated = true;
        activeUser = usernameInput;
        activeRole = uData.role;
        activeFullName = uData.full_name || usernameInput;
        
        // Guardar sesión temporal en sessionStorage
        sessionStorage.setItem("simq_session", JSON.stringify({
            authenticated: true,
            activeUser: activeUser,
            activeRole: activeRole,
            activeFullName: activeFullName
        }));
        
        // Ocultar login y mostrar layout principal
        document.getElementById('loginContainer').style.display = 'none';
        document.getElementById('appLayout').style.display = 'flex';
        
        updateUserBadge();
        switchTab('dashboard');
        translateStaticUI();
    } else {
        errorDiv.innerText = currentLang === 'es' ? '❌ Usuario o contraseña incorrectos.' : '❌ Incorrect username or password.';
        errorDiv.style.display = 'block';
    }
}

// Cerrar sesión (Action Handler)
function handleLogout() {
    authenticated = false;
    activeUser = '';
    activeRole = '';
    activeFullName = '';
    
    sessionStorage.removeItem("simq_session");
    
    document.getElementById('appLayout').style.display = 'none';
    document.getElementById('loginContainer').style.display = 'flex';
    
    document.getElementById('usernameInput').value = '';
    document.getElementById('passwordInput').value = '';
    document.getElementById('loginErrorMsg').style.display = 'none';
}

// Actualizar la tarjeta de perfil de usuario en la barra lateral
function updateUserBadge() {
    document.getElementById('userFullName').innerText = activeFullName;
    const roleLabel = activeRole === 'admin' 
        ? (currentLang === 'es' ? 'Administrador' : 'Administrator') 
        : (currentLang === 'es' ? 'Cliente / Visor' : 'Client / Viewer');
    document.getElementById('userRole').innerText = `${roleLabel} (${activeUser})`;
    
    const adminBtn = document.getElementById('nav-admin');
    if (adminBtn) {
        if (activeRole === 'admin') {
            adminBtn.style.display = 'flex';
        } else {
            adminBtn.style.display = 'none';
        }
    }
}

// Verificar sesión existente al cargar la página
function checkSession() {
    const session = sessionStorage.getItem("simq_session");
    if (session) {
        try {
            const data = JSON.parse(session);
            if (data.authenticated) {
                authenticated = true;
                activeUser = data.activeUser;
                activeRole = data.activeRole;
                activeFullName = data.activeFullName;
                
                document.getElementById('loginContainer').style.display = 'none';
                document.getElementById('appLayout').style.display = 'flex';
                updateUserBadge();
                switchTab('dashboard');
                return;
            }
        } catch (e) {
            console.error("Error al procesar la sesión recuperada:", e);
        }
    }
    document.getElementById('loginContainer').style.display = 'flex';
    document.getElementById('appLayout').style.display = 'none';
}

// ══════════════════════════════════════════════════════════════════
// PERSISTENCIA DE ANÁLISIS (JSON EXPORT/IMPORT)
// ══════════════════════════════════════════════════════════════════

// Exportar análisis en un archivo .json y descargarlo mediante navegador
function downloadAnalysisJson() {
    const data = {
        app: "SIM-Q RELIARISK",
        version: "1.0",
        timestamp: new Date().toISOString(),
        platformsData: platformsData,
        userMitigations: userMitigations
    };
    
    const jsonStr = JSON.stringify(data, null, 2);
    const blob = new Blob([jsonStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = `Analisis_SIM_Q_Reliarisk_${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    const status = document.getElementById('jsonLoadStatus');
    if (status) {
        status.innerHTML = `<span style="color: #10b981; font-weight: 600;">✅ ${currentLang === 'es' ? 'Descargado con éxito' : 'Downloaded successfully'}</span>`;
        setTimeout(() => { status.innerHTML = ''; }, 3000);
    }
}

// Guardar simulación activa localmente en LocalStorage
function saveAnalysisJson() {
    localStorage.setItem("simq_active_analysis", JSON.stringify({
        platformsData: platformsData,
        userMitigations: userMitigations
    }));
    
    const status = document.getElementById('jsonLoadStatus');
    if (status) {
        status.innerHTML = `<span style="color: #10b981; font-weight: 600;">💾 ${currentLang === 'es' ? 'Guardado local completado' : 'Saved locally'}</span>`;
        setTimeout(() => { status.innerHTML = ''; }, 3000);
    }
}

// Desencadenar la subida del archivo JSON
function triggerJsonUpload() {
    document.getElementById('jsonFileUploader').click();
}

// Leer archivo JSON subido y restaurar la sesión
function loadAnalysisJson(event) {
    const file = event.target.files[0];
    const status = document.getElementById('jsonLoadStatus');
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = JSON.parse(e.target.result);
            if (data.app !== "SIM-Q RELIARISK") {
                status.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ ${currentLang === 'es' ? 'Archivo no compatible con SIM-Q' : 'File not compatible with SIM-Q'}</span>`;
                return;
            }
            
            if (data.platformsData) {
                platformsData = data.platformsData;
            }
            if (data.userMitigations) {
                userMitigations = data.userMitigations;
            }
            
            // Guardar estado restaurado en LocalStorage
            saveAnalysisJson();
            
            status.innerHTML = `<span style="color: #10b981; font-weight: 600;">✅ ${currentLang === 'es' ? 'Análisis cargado con éxito' : 'Analysis loaded successfully'}</span>`;
            switchTab(currentTab);
        } catch (err) {
            status.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ ${currentLang === 'es' ? 'Error al leer el archivo JSON' : 'Error reading JSON file'}</span>`;
            console.error(err);
        }
    };
    reader.readAsText(file);
}

// ══════════════════════════════════════════════════════════════════
// PANEL DE ADMINISTRACIÓN DE USUARIOS (CRUD & CONFIG)
// ══════════════════════════════════════════════════════════════════

let activeAdminTab = 'create-user';

function switchAdminTab(tabId) {
    activeAdminTab = tabId;
    renderAdminPanel();
}

// Renderizar panel de administración de usuarios en el DOM
function renderAdminPanel() {
    const mainContent = document.getElementById('mainContent');
    mainContent.innerHTML = '';
    
    const isES = currentLang === 'es';
    
    const adminView = document.createElement('div');
    adminView.className = 'admin-panel-view';
    
    const tabCreateActive = activeAdminTab === 'create-user' ? 'active' : '';
    const tabListActive = activeAdminTab === 'users-list' ? 'active' : '';
    const tabConfigActive = activeAdminTab === 'config-toml' ? 'active' : '';
    
    adminView.innerHTML = `
        <div class="admin-header">
            <h2>👥 ${isES ? 'Gestión de Usuarios y Accesos' : 'User and Access Management'}</h2>
            <p class="admin-subtitle">
                ${isES ? 'Administre las credenciales de los clientes y usuarios que pueden acceder al sistema de SIM-Q RELIARISK.' : 'Manage credentials for clients and users who can access the SIM-Q RELIARISK system.'}
            </p>
        </div>
        <div class="admin-tabs">
            <button class="admin-tab-btn ${tabCreateActive}" onclick="switchAdminTab('create-user')">➕ ${isES ? 'Crear Nuevo Usuario' : 'Create New User'}</button>
            <button class="admin-tab-btn ${tabListActive}" onclick="switchAdminTab('users-list')">📋 ${isES ? 'Lista y Edición de Usuarios' : 'Users List & Editing'}</button>
            <button class="admin-tab-btn ${tabConfigActive}" onclick="switchAdminTab('config-toml')">⚙️ ${isES ? 'Configuración Secrets Cloud' : 'Streamlit Cloud Secrets Config'}</button>
        </div>
        <div class="admin-tab-content" id="adminTabContent"></div>
    `;
    
    mainContent.appendChild(adminView);
    
    const tabContent = document.getElementById('adminTabContent');
    const users = loadUsers();
    
    if (activeAdminTab === 'create-user') {
        tabContent.innerHTML = `
            <div class="info-banner">
                ${isES ? 'Genere una cuenta con contraseña cifrada (PBKDF2-HMAC-SHA256) para un cliente o miembro del equipo.' : 'Generate an account with a secure hashed password (PBKDF2-HMAC-SHA256) for a client or team member.'}
            </div>
            <form onsubmit="handleCreateUser(event)" style="display: flex; flex-direction: column; gap: 16px;">
                <div class="form-row">
                    <div class="form-group">
                        <label>${isES ? 'Nombre de Usuario (ID de Acceso)' : 'Username (Access ID)'}</label>
                        <input type="text" id="newUsername" placeholder="${isES ? 'ej. cliente_ekbalam' : 'e.g. client_ekbalam'}" required autocomplete="off">
                    </div>
                    <div class="form-group">
                        <label>${isES ? 'Nombre Completo / Empresa' : 'Full Name / Company'}</label>
                        <input type="text" id="newFullname" placeholder="${isES ? 'ej. Pemex Exploración y Producción' : 'e.g. Pemex Exploration'}" required autocomplete="off">
                    </div>
                </div>
                <div class="form-row">
                    <div class="form-group">
                        <label>${isES ? 'Contraseña Inicial' : 'Initial Password'}</label>
                        <input type="password" id="newPassword" placeholder="${isES ? 'Mínimo 4 caracteres' : 'Min 4 characters'}" required autocomplete="new-password">
                    </div>
                    <div class="form-group">
                        <label>${isES ? 'Rol de Acceso' : 'Access Role'}</label>
                        <select id="newRole" style="width: 100%;">
                            <option value="client">${isES ? '👤 Cliente / Visor (Solo consulta y análisis)' : '👤 Client / Viewer (Query and analysis only)'}</option>
                            <option value="admin">${isES ? '👑 Administrador (Acceso total y gestión)' : '👑 Administrator (Full access and management)'}</option>
                        </select>
                    </div>
                </div>
                <div class="form-admin-actions">
                    <button type="submit" class="btn-primary" style="width: 100%; max-width: 250px;">✨ ${isES ? 'Crear Usuario' : 'Create User'}</button>
                </div>
                <div id="createUserResult" class="load-status-msg" style="margin-top: 8px;"></div>
            </form>
        `;
    } else if (activeAdminTab === 'users-list') {
        const trs = Object.keys(users).map(u => {
            const uData = users[u];
            const roleLabel = uData.role === 'admin' ? `👑 ${isES ? 'Administrador' : 'Admin'}` : `👤 ${isES ? 'Cliente' : 'Client'}`;
            const truncHash = uData.hash ? uData.hash.slice(0, 16) + '...' : '-';
            return `
                <tr>
                    <td><strong>${u}</strong></td>
                    <td>${uData.full_name || u}</td>
                    <td>${roleLabel}</td>
                    <td><code>${truncHash}</code></td>
                </tr>
            `;
        }).join('');
        
        const selectOptions = Object.keys(users).map(u => `<option value="${u}">${u}</option>`).join('');
        
        tabContent.innerHTML = `
            <h3>${isES ? 'Usuarios Registrados en el Sistema' : 'Users Registered in the System'}</h3>
            <div style="overflow-x: auto; margin-top: 12px; margin-bottom: 24px;">
                <table class="admin-users-table">
                    <thead>
                        <tr>
                            <th>${isES ? 'Usuario (ID)' : 'Username (ID)'}</th>
                            <th>${isES ? 'Nombre / Organización' : 'Name / Organization'}</th>
                            <th>${isES ? 'Rol' : 'Role'}</th>
                            <th>${isES ? 'Hash Cifrado (PBKDF2)' : 'Hashed (PBKDF2)'}</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${trs}
                    </tbody>
                </table>
            </div>
            
            <div class="admin-manage-grid">
                <div class="admin-card-section">
                    <h4>🔑 ${isES ? 'Cambiar Contraseña de Usuario' : 'Change User Password'}</h4>
                    <form onsubmit="handleChangePassword(event)" style="display: flex; flex-direction: column; gap: 12px;">
                        <div class="form-group">
                            <label>${isES ? 'Seleccionar Usuario' : 'Select User'}</label>
                            <select id="pwdUserSelect">
                                ${selectOptions}
                            </select>
                        </div>
                        <div class="form-group">
                            <label>${isES ? 'Nueva Contraseña' : 'New Password'}</label>
                            <input type="password" id="pwdNewPassword" required autocomplete="new-password">
                        </div>
                        <button type="submit" class="btn-primary" style="margin-top: 8px;">${isES ? 'Actualizar Contraseña' : 'Update Password'}</button>
                        <div id="changePwdResult" class="load-status-msg"></div>
                    </form>
                </div>
                
                <div class="admin-card-section">
                    <h4>🗑️ ${isES ? 'Eliminar Usuario' : 'Delete User'}</h4>
                    <form onsubmit="handleDeleteUser(event)" style="display: flex; flex-direction: column; gap: 12px;">
                        <div class="form-group">
                            <label>${isES ? 'Seleccionar Usuario a Eliminar' : 'Select User to Delete'}</label>
                            <select id="delUserSelect">
                                ${selectOptions}
                            </select>
                        </div>
                        <button type="submit" class="btn-login" style="background: var(--risk-high); border-color: transparent; margin-top: 8px;">❌ ${isES ? 'Eliminar Usuario' : 'Delete User'}</button>
                        <div id="deleteUserResult" class="load-status-msg"></div>
                    </form>
                </div>
            </div>
        `;
    } else if (activeAdminTab === 'config-toml') {
        let tomlStr = "";
        for (const [uName, uData] of Object.entries(users)) {
            tomlStr += `[users.${uName}]\n`;
            tomlStr += `salt = "${uData.salt}"\n`;
            tomlStr += `hash = "${uData.hash}"\n`;
            tomlStr += `role = "${uData.role}"\n`;
            tomlStr += `full_name = "${uData.full_name}"\n\n`;
        }
        
        tabContent.innerHTML = `
            <h3>${isES ? 'Configuración Compartida de Secrets (Streamlit Cloud)' : 'Shared Secrets Config (Streamlit Cloud)'}</h3>
            <p style="font-size: 13px; color: var(--text-muted); margin-top: 8px; margin-bottom: 16px; line-height: 1.4;">
                ${isES 
                  ? 'Para que la suite de herramientas trabaje de forma unificada, puedes copiar este bloque TOML de usuarios cifrados y pegarlo directamente en la sección <strong>Secrets</strong> de la consola de administración de <strong>RBI-Q RELIARISK</strong> en Streamlit Cloud. De esta forma, ambas aplicaciones compartirán las mismas credenciales.'
                  : 'For the suite of tools to work in a unified way, you can copy this TOML block of hashed users and paste it directly in the <strong>Secrets</strong> section of the <strong>RBI-Q RELIARISK</strong> admin console in Streamlit Cloud. This way, both apps will share the same credentials.'
                }
            </p>
            <div style="position: relative;">
                <pre style="background: var(--surface-container-low); padding: 16px; border-radius: 8px; font-family: monospace; font-size: 11px; overflow-x: auto; outline: 1px solid var(--outline-variant); max-height: 350px;">${tomlStr.trim()}</pre>
                <button onclick="navigator.clipboard.writeText(\`${tomlStr.trim().replace(/`/g, '\\`').replace(/\$/g, '\\$')}\`).then(() => { alert('${isES ? 'Copiado al portapapeles' : 'Copied to clipboard'}'); })" class="btn-sm" style="position: absolute; top: 10px; right: 10px; background: #ffffff;">📋 ${isES ? 'Copiar' : 'Copy'}</button>
            </div>
            <p style="font-size: 11px; color: var(--text-muted); margin-top: 8px;">
                🔒 ${isES ? 'El cifrado PBKDF2-HMAC-SHA256 con 100,000 iteraciones garantiza que las contraseñas no puedan descifrarse de forma directa.' : 'PBKDF2-HMAC-SHA256 hashing with 100,000 iterations ensures that passwords cannot be decrypted directly.'}
            </p>
        `;
    }
}

// Crear un usuario nuevo (Admin Handler)
async function handleCreateUser(e) {
    e.preventDefault();
    const isES = currentLang === 'es';
    const resultDiv = document.getElementById('createUserResult');
    
    const username = document.getElementById('newUsername').value.trim().toLowerCase();
    const fullname = document.getElementById('newFullname').value.trim();
    const password = document.getElementById('newPassword').value;
    const role = document.getElementById('newRole').value;
    
    if (password.length < 4) {
        resultDiv.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ ${isES ? 'La contraseña debe tener al menos 4 caracteres.' : 'Password must be at least 4 characters long.'}</span>`;
        return;
    }
    
    const users = loadUsers();
    if (users[username]) {
        resultDiv.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ ${isES ? `El usuario '${username}' ya existe.` : `User '${username}' already exists.`}</span>`;
        return;
    }
    
    resultDiv.innerHTML = `<span style="color: var(--primary); font-weight: 600;">⏳ ${isES ? 'Cifrando contraseña...' : 'Hashing password...'}</span>`;
    
    try {
        const credentials = await hashPassword(password);
        users[username] = {
            salt: credentials.salt,
            hash: credentials.hash,
            role: role,
            full_name: fullname
        };
        
        saveUsers(users);
        resultDiv.innerHTML = `<span style="color: #10b981; font-weight: 600;">✅ ${isES ? `Usuario '${username}' creado con éxito.` : `User '${username}' successfully created.`}</span>`;
        
        setTimeout(() => {
            switchAdminTab('users-list');
        }, 1200);
    } catch(err) {
        resultDiv.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ Error: ${err.message}</span>`;
    }
}

// Modificar contraseña de usuario (Admin Handler)
async function handleChangePassword(e) {
    e.preventDefault();
    const isES = currentLang === 'es';
    const resultDiv = document.getElementById('changePwdResult');
    
    const username = document.getElementById('pwdUserSelect').value;
    const newPassword = document.getElementById('pwdNewPassword').value;
    
    if (newPassword.length < 4) {
        resultDiv.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ ${isES ? 'La contraseña debe tener al menos 4 caracteres.' : 'Password must be at least 4 characters long.'}</span>`;
        return;
    }
    
    resultDiv.innerHTML = `<span style="color: var(--primary); font-weight: 600;">⏳ ${isES ? 'Cifrando nueva contraseña...' : 'Hashing new password...'}</span>`;
    
    try {
        const credentials = await hashPassword(newPassword);
        const users = loadUsers();
        if (users[username]) {
            users[username].salt = credentials.salt;
            users[username].hash = credentials.hash;
            saveUsers(users);
            resultDiv.innerHTML = `<span style="color: #10b981; font-weight: 600;">✅ ${isES ? `Contraseña actualizada para '${username}'.` : `Password updated for '${username}'.`}</span>`;
        }
        
        setTimeout(() => {
            switchAdminTab('users-list');
        }, 1200);
    } catch(err) {
        resultDiv.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ Error: ${err.message}</span>`;
    }
}

// Eliminar un usuario (Admin Handler)
function handleDeleteUser(e) {
    e.preventDefault();
    const isES = currentLang === 'es';
    const resultDiv = document.getElementById('deleteUserResult');
    
    const username = document.getElementById('delUserSelect').value;
    
    if (username === activeUser) {
        resultDiv.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ ${isES ? 'No puedes eliminar tu propio usuario en uso.' : 'You cannot delete your own active user account.'}</span>`;
        return;
    }
    
    const users = loadUsers();
    
    // Verificar que quede al menos un administrador
    const adminsCount = Object.keys(users).filter(u => users[u].role === 'admin').length;
    if (users[username].role === 'admin' && adminsCount <= 1) {
        resultDiv.innerHTML = `<span style="color: #ef4444; font-weight: 600;">❌ ${isES ? 'No se puede eliminar el único administrador del sistema.' : 'Cannot delete the only administrator in the system.'}</span>`;
        return;
    }
    
    delete users[username];
    saveUsers(users);
    
    resultDiv.innerHTML = `<span style="color: #10b981; font-weight: 600;">✅ ${isES ? `Usuario '${username}' eliminado.` : `User '${username}' deleted.`}</span>`;
    
    setTimeout(() => {
        switchAdminTab('users-list');
    }, 1200);
}



