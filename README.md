# PYXCEL

Aplicación de hoja de cálculo estilo Python IDLE escrita en Python con PySide6. Presentamos un diseño oscuro profesional inspirado en Catppuccin Mocha.

## Características

- **Interfaz moderna**: Tema oscuro profesional con acentos en azul
- **Pantalla de carga**: Splash screen animado con barra de progreso
- **Toolbar con iconos**: Botones intuitivos con tooltips informativos
- **Terminal interactiva**: Ejecuta código Python directamente con resaltado de sintaxis y variables persistentes
- **Hoja de cálculo**: Gestión de datos en formato tabular con soporte para múltiples hojas
- **Fórmulas avanzadas**: Motor completo compatible con Excel (VLOOKUP, INDEX, MATCH, XLOOKUP, FILTER, etc.)
- **Formato de celdas visual**: Negrita, cursiva, subrayado, color de fuente, color de fondo, tamaño
- **Gráficos**: Barras, líneas, circular, dispersión, área con tema oscuro
- **Ordenamiento multinivel**: Ordenar por múltiples columnas
- **Filtrado**: Filtrar datos por condiciones
- **Macros**: Sistema de macros con persistencia en archivos
- **Operaciones**: Copiar, pegar, cortar, deshacer, rehacer
- **Buscar y reemplazar**: Funcionalidad completa de búsqueda
- **Importar/Exportar**: Archivos Excel (.xlsx), CSV
- **Multi-plataforma**: Windows, Mac, Linux

## Requisitos

- Python 3.13+
- PySide6 6.8.0
- pandas, numpy, openpyxl
- matplotlib, seaborn
- formulas

## Instalación

```bash
cd PYXCEL
uv sync
```

## Uso

```bash
python main.py
```

## Estructura del Proyecto

```
PYXCEL/
├── main.py                 # Punto de entrada
├── pyproject.toml          # Configuración del proyecto
├── README.md              # Este archivo
├── pyxcel/
│   ├── app.py             # Ventana principal y terminal interactiva
│   ├── models/
│   │   ├── workbook.py    # Gestor de hojas de cálculo
│   │   └── spreadsheet.py # Modelo de datos (celdas, formatos, fórmulas)
│   ├── ui/
│   │   ├── table.py       # Widget de tabla
│   │   ├── toolbar.py     # Barra de herramientas y menús
│   │   ├── dialogs.py     # Diálogos (buscar, reemplazar, formato, gráficos)
│   │   └── theme.py       # Sistema de temas y estilos
│   ├── utils/
│   │   ├── file_handler.py   # Manejo de archivos
│   │   └── chart_builder.py # Constructor de gráficos
│   ├── engine/
│   │   └── formulas.py    # Motor de fórmulas
│   └── macros/
│       └── macro_system.py # Sistema de macros
```

## Diseño Visual

### Tema Oscuro Catppuccin

PYXCEL implementa un tema oscuro profesional inspirado en Catppuccin Mocha:

| Elemento | Color |
|----------|-------|
| Fondo principal | `#1E1E2E` |
| Fondo de paneles | `#181825` |
| Acento primario | `#89B4FA` |
| Éxito | `#A6E3A1` |
| Error | `#F38BA8` |
| Warning | `#F9E2AF` |
| Texto | `#CDD6F4` |

### Tipografía

- **Spreadsheet**: JetBrains Mono (monoespaciada)
- **Terminal**: JetBrains Mono / Consolas
- **UI**: Segoe UI

## Descripción de Componentes

### app.py
Contiene la clase principal `PYXCEL`, `TerminalWidget` y splash screen:
- Pantalla de carga animada con barra de progreso
- Interfaz con splitters y paneles
- Terminal interactiva con salida codificada por colores y variables persistentes
- Gestión de hojas de cálculo con tabs
- Sistema de gráficos integrado

### ui/theme.py
Sistema de temas profesionales:
- `ThemeColors`: Paleta de colores Catppuccin
- `ThemeFonts`: Configuración de fuentes
- `get_app_stylesheet()`: Stylesheet global
- `get_palette()`: Paleta Qt

### models/spreadsheet.py
- `SpreadsheetModel`: Modelo Qt para datos tabulares
- `CellData`: Datos de cada celda (valor, fórmula, formato)
- `CellFormat`: Formato de celda (fuente, color, alineación, etc.)

### ui/toolbar.py
- `ToolbarManager`: Barra de herramientas con iconos
- `MenuManager`: Menús en español (Archivo, Editar, Shell, Depurar, Opciones, Ventana, Ayuda)

### engine/formulas.py
Motor de evaluación de fórmulas compatible con Excel con funciones avanzadas:
- **Búsqueda**: VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH
- **Matemáticas**: SUMIF, COUNTIF, AVERAGEIF, SUM, AVERAGE, COUNT
- **Manipulación**: FILTER, UNIQUE, SORT
- **Lógicas**: IF, AND, OR, NOT

### macros/macro_system.py
Sistema de macros con persistencia:
- Registro de macros con decoradores
- Guardado/carga de macros en archivos
- Exportación de macros a archivos Excel

## Gráficos

PYXCEL soporta los siguientes tipos de gráficos:
- Barras
- Barras apiladas
- Líneas
- Líneas apiladas
- Circular (Pie)
- Dispersión (Scatter)
- Área
- Área apilada

Los gráficos se muestran en una ventana emergente con el tema oscuro aplicado.

## Pantalla de Carga

Al iniciar la aplicación se muestra una splash screen con:
- Logo "PYXCEL" con diseño profesional
- Barra de progreso animada
- Mensajes de estado durante la carga
- Tema oscuro aplicado

## Interfaz

La aplicación tiene dos paneles principales:

1. **Panel Izquierdo**:
   - Pestaña "Shell": Historial de comandos ejecutados
   - Pestaña "Archivos": Árbol de hojas del libro

2. **Panel Derecho**:
   - Pestaña "Terminal": Terminal interactiva
   - Pestaña "Editor": Hojas de cálculo con tabs

## Terminal Interactiva

La terminal permite ejecutar código Python directamente con variables persistentes:
- Escribe comandos y presiona Enter para ejecutar
- Usa flechas arriba/abajo para navegar el historial
- Las variables persisten entre ejecuciones
- Colores para diferentes tipos de salida:
  - Blanco (`#CDD6F4`): Salida de print
  - Mauve (`#CBA6F7`): Valores de retorno
  - Rojo (`#F38BA8`): Errores
- Bordes decorativos estilo IDE

Ejemplos:
```python
a = 2
b = 3
print(a + b)  # Salida: 5
a * b         # Retorna: 6
for i in range(5):
    print(i)
```

## Atajos de Teclado

- `Ctrl+N`: Nuevo archivo
- `Ctrl+O`: Abrir archivo
- `Ctrl+S`: Guardar
- `Ctrl+Shift+S`: Guardar como
- `Ctrl+Z`: Deshacer
- `Ctrl+Y`: Rehacer
- `Ctrl+X`: Cortar
- `Ctrl+C`: Copiar
- `Ctrl+V`: Pegar
- `Ctrl+F`: Buscar
- `Ctrl+H`: Reemplazar
- `F5`: Ejecutar (Depurar)
- `Enter`: Ejecutar celda (Terminal)

## Macros

PYXCEL incluye un sistema de macros que permite automatizar tareas:

```python
from pyxcel.macros.macro_system import macro

@macro(name="mi_macro", description="Mi primera macro")
def mi_macro():
    # Código de la macro
    pass
```

Las macros se pueden guardar y cargar con los archivos Excel.

## Licencia

MIT License

## Autor

Edilberto Madrigal

## Testing

PYXCEL incluye una suite de pruebas completa con 68 tests:

```bash
# Ejecutar todos los tests
python -m pytest tests/ -v

# Ejecutar tests específicos
python -m pytest tests/test_formulas.py -v
python -m pytest tests/test_workbook.py -v
python -m pytest tests/test_file_handler.py -v
python -m pytest tests/test_spreadsheet.py -v
```

### Cobertura de Tests

| Módulo | Tests | Descripción |
|--------|-------|-------------|
| test_formulas.py | 46 | Motor de fórmulas (SUM, VLOOKUP, INDEX, etc.) |
| test_file_handler.py | 6 | Persistencia (CSV, XLSX) |
| test_workbook.py | 13 | Gestión de hojas |
| test_spreadsheet.py | 13 | Modelo de datos |

## Contribuir

1. Fork el repositorio
2. Crea una rama (`git checkout -b feature/nueva-caracteristica`)
3. Commit tus cambios (`git commit -am 'Agrega nueva característica'`)
4. Push a la rama (`git push origin feature/nueva-caracteristica`)
5. Crea un Pull Request
