# PYXCEL

Aplicación de hoja de cálculo estilo Python IDLE escrita en Python con PySide6.

## Características

- **Interfaz estilo IDLE**: Panel izquierdo con Terminal/Explorador y panel derecho con Terminal/Editor
- **Terminal interactiva**: Ejecuta código Python directamente en la aplicación (estilo MS-DOS)
- **Hoja de cálculo**: Gestión de datos en formato tabular con soporte para múltiples hojas
- **Fórmulas**: Motor de fórmulas compatible con Excel
- **Formato de celdas**: Fuente, tamaño, negrita, cursiva, subrayado, color, alineación
- **Operaciones**: Copiar, pegar, cortar, deshacer, rehacer
- **Buscar y reemplazar**: Funcionalidad completa de búsqueda
- **Ordenamiento**: Ordenar datos por columnas
- **Importar/Exportar**: Archivos Excel (.xlsx), CSV

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
│   │   └── dialogs.py     # Diálogos (buscar, reemplazar, formato, gráficos)
│   ├── utils/
│   │   ├── file_handler.py   # Manejo de archivos
│   │   └── chart_builder.py # Constructor de gráficos
│   ├── engine/
│   │   └── formulas.py    # Motor de fórmulas
│   └── macros/
│       └── macro_system.py # Sistema de macros
```

## Descripción de Componentes

### app.py
Contiene la clase principal `PYXCEL` y `TerminalWidget`:
- Interfaz con splitters y paneles
- Terminal interactiva estilo MS-DOS (fondo negro, texto verde)
- Gestión de hojas de cálculo

### models/spreadsheet.py
- `SpreadsheetModel`: Modelo Qt para datos tabulares
- `CellData`: Datos de cada celda (valor, fórmula, formato)
- `CellFormat`: Formato de celda (fuente, color, alineación, etc.)

### ui/toolbar.py
- `ToolbarManager`: Barra de herramientas
- `MenuManager`: Menús en español (Archivo, Editar, Shell, Depurar, Opciones, Ventana, Ayuda)

### engine/formulas.py
Motor de evaluación de fórmulas compatible con Excel.

## Interfaz

La aplicación tiene dos paneles principales:

1. **Panel Izquierdo**:
   - Pestaña "Terminal": Lista de comandos ejecutados
   - Pestaña "Explorador": Árbol de hojas del libro

2. **Panel Derecho**:
   - Pestaña "Terminal": Terminal interactiva estilo MS-DOS
   - Pestaña "Editor": Hojas de cálculo

## Terminal Interactiva

La terminal permite ejecutar código Python directamente:
- Escribe comandos y presiona Enter para ejecutar
- Usa flechas arriba/abajo para navegar el historial
- Fondo negro con texto verde estilo MS-DOS
- Muestra errores en rojo

Ejemplos:
```python
2 + 2
print("Hola mundo")
import math
math.sqrt(16)
```

## Atajos de Teclado

- `Ctrl+N`: Nuevo archivo
- `Ctrl+O`: Abrir archivo
- `Ctrl+G`: Guardar
- `Ctrl+Shift+S`: Guardar como
- `Ctrl+Z`: Deshacer
- `Ctrl+Y`: Rehacer
- `Ctrl+X`: Cortar
- `Ctrl+C`: Copiar
- `Ctrl+V`: Pegar
- `Ctrl+F`: Buscar
- `Ctrl+H`: Reemplazar

## Licencia

MIT License

## Autor

Edilberto Madrigal
