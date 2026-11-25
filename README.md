# Flow Matching & Rectified Flows — material de la presentación

Presentación de 45 minutos para la asignatura de Aprendizaje Profundo, en grupo de dos.

## Contenido

| Ruta | Qué es |
|---|---|
| `slides/FlowMatching_RectifiedFlows.pptx` | Diapositivas (29) con notas del presentador, reparto P1/P2 y tiempos |
| `notebooks/01_flow_matching_2d.ipynb` | Flow Matching desde cero en 2D: entrenamiento, campo de velocidades, calidad vs. NFE, conexión con difusión (score, SDE) |
| `notebooks/02_rectified_flow_reflow.ipynb` | Rectified Flow: cruces vs. reflow, OT-CFM, generación en 1–4 pasos, destilación |
| `notebooks/figuras/` | Figuras generadas por los notebooks (las usan las diapositivas) |
| `slides/build_deck.py` | Genera el .pptx (python-pptx) |
| `slides/extra_figs.py` | Ecuaciones (mathtext), figura de schedules y portada |
| `slides/render.ps1` | Exporta las diapositivas a PNG con PowerPoint (revisión visual) |

## Reparto y tiempos (45 min)

| Bloque | Diapositivas | Quién | Tiempo |
|---|---|---|---|
| 1. Motivación y flujos continuos | 1–6 | P1 | ~8.5 min |
| 2. Flow Matching | 7–13 | P1 | ~11.5 min |
| 3. Conexión con difusión | 14–17 | P2 | ~6 min |
| 4. Rectified Flow | 18–24 | P2 | ~10 min |
| 5. Demo (notebooks) | 25 | P2 | ~5 min |
| 6. Estado del arte y conclusiones | 26–29 | P2 | ~3.5 min |

Los minutos exactos de cada diapositiva están en sus notas.

## Ejecutar los notebooks

- **Colab**: subir el `.ipynb` y *Ejecutar todo*. Funcionan en CPU: ~1 min el notebook 1 y ~3 min el 2.
- **Local**: `python -m venv .venv`, luego `.venv\Scripts\pip install torch numpy matplotlib scipy jupyter` y abrir con Jupyter.

## Regenerar las diapositivas

```powershell
cd slides
..\.venv\Scripts\python extra_figs.py
..\.venv\Scripts\python build_deck.py
```

Si cambian los resultados del notebook 2, actualizar el diccionario `RES` de `build_deck.py` (diapositiva 23) con la tabla que imprime la sección 3.
