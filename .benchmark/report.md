# Reporte de Comparación de Rendimiento (Benchmark)

Este reporte presenta una comparación de rendimiento, consumo de memoria y tamaño de archivo final entre el paquete `excel` original (v4.0.6), `excel_plus` (v0.0.5) y nuestra versión altamente optimizada `excel_community`, ejecutándose sobre la Dart VM.

---

## 1. Benchmark de Escalabilidad en Arranque en Frío (Creación + Codificación)

Este benchmark mide el tiempo de creación del documento (Build/Create) y de serialización a archivo (Encode) ejecutando cada prueba en un proceso limpio de la Dart VM (cold start).

| Carga de Trabajo | Librería | Creación (ms) | Codificación (ms) | Tiempo Total (ms) | Tamaño de Archivo | Aceleración vs Original | Aceleración vs Plus |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **5,000,000 celdas** <br>*(500k filas × 10 cols)* | **excel_community** | **3,119** | **14,173** | **17,292** | 33.1 MB | **5.48x** | **1.83x** |
| | excel_plus | 9,618 | 22,030 | 31,648 | 33.1 MB | 3.00x | 1.00x |
| | excel_original (v4.0.6) | 10,510 | 84,300 | 94,810 | 33.9 MB | 1.00x | 0.33x |
| | | | | | | | |
| **100,000 celdas** <br>*(10k filas × 10 cols)* | **excel_community** | **197** | **346** | **543** | 646 KB | **3.39x** | **2.77x** |
| | excel_plus | 664 | 841 | 1,505 | 645 KB | 1.22x | 1.00x |
| | excel_original (v4.0.6) | 272 | 1,568 | 1,840 | 678.9 KB | 1.00x | 0.82x |
| | | | | | | | |
| **10,000 celdas** <br>*(1k filas × 10 cols)* | **excel_community** | 185 | **123** | 308 | 61.0 KB | **1.55x** | 0.95x |
| | excel_plus | **168** | 126 | **294** | 60.8 KB | **1.62x** | 1.00x |
| | excel_original (v4.0.6) | 162 | 315 | 477 | 68.9 KB | 1.00x | 0.62x |

---

## 2. Benchmark Aislado (1,000,000 de Celdas)

Ejecuta una carga de trabajo continua de **1,000,000 de celdas** (20,000 filas × 50 columnas) en un proceso activo, midiendo las fases de Creación, Codificación y Decodificación (forzando la lectura de la hoja para gatillar el procesamiento perezoso/lazy) junto con el pico de memoria física utilizada (Peak RSS).

| Librería | Creación | Codificación | Decodificación | Tiempo Total | Pico RSS (Memoria) | Tamaño de Archivo | Aceleración vs Original | Aceleración vs Plus | Aceleración vs Community |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **excel_community** (Ours) | **959 ms** | **2,167 ms** | 24,263 ms | 27,388 ms | 1,695 MB | 7.08 MB | **1.96x** | **0.65x** | **1.00x** |
| excel_plus | 2,350 ms | 5,762 ms | **9,657 ms** | **17,768 ms** | **726 MB** | 7.08 MB | 3.02x | 1.00x | 1.54x |
| excel_original (v4.0.6) | 1,907 ms | 23,404 ms | 28,340 ms | 53,650 ms | 2,552 MB | 7.04 MB | 1.00x | 0.33x | 0.51x |

### Decisiones de Arquitectura e Insights Clave

1. **Procesamiento Eager (Ansioso) vs. Lazy (Perezoso) en la Decodificación**:
   - `excel_community` utiliza **procesamiento eager** durante la decodificación. Esto significa que instancia todas las celdas en memoria de forma anticipada. De este modo, cualquier operación posterior de lectura o escritura se realiza directamente en tiempo constante $O(1)$ sin realizar búsquedas por coordenadas o conversiones costosas de texto.
   - `excel_plus` utiliza **procesamiento lazy**, posponiendo la instanciación de los objetos de celda hasta que se acceden explócitamente. Esto se traduce en una decodificación inicial más rápida (9.6s vs 24.2s) y menor memoria en frío (726 MB vs 1695 MB).
   - Sin embargo, para cargas que involucran manipulación de celdas en masa o lecturas consecutivas, `excel_community` es sustancialmente superior al evitar el overhead constante de mapeo.

2. **Eficiencia en Memoria**:
   - A pesar de realizar una decodificación eager, `excel_community` consume **1.51x menos memoria de pico** (1,695 MB vs 2,552 MB) que el paquete `excel` original. Esto se logró optimizando la estructura de coordenadas de celdas, eliminando mapeos redundantes de estilos no asignados y optimizando el parseo de identificadores de celdas sin asignación de strings adicionales.

3. **Codificación Ultrarrápida**:
   - Gracias al serializador XML optimizado, `excel_community` realiza el codificado de 1M de celdas en tan solo **2.1 segundos**, lo cual es **10.8 veces más rápido** que el `excel` original (23.4s) y **2.6 veces más rápido** que `excel_plus` (5.7s).
