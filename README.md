## Archivos

- `entrada/`: Excels a compararse.
- `src/analizar_internet.py`: script del análisis.
- `salida/`: aquí se genera `analisis_internet.xlsx`.

## Preparar y ejecutar

Instala Python 3.10 o posterior. En Windows, abre PowerShell en esta carpeta y ejecuta:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe .\src\analizar_internet.py
```

El nombre se normaliza quitando el sufijo final `W` seguido de números (por ejemplo, `EBAUTISTAW111` → `EBAUTISTA`). La búsqueda en SGA compara el login sin distinguir mayúsculas.

La comparación considera coincidentes el rol de SGA y el código convertido cuando sus valores son iguales; por ejemplo, `NSC` coincide con `NSC`. En el reporte, “Valor esperado” es el rol de SGA y “Valor encontrado” el código del grupo. También se informa si coincide, no coincide, falta el usuario o no existe código en el catálogo.
