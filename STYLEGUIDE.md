# Guía de Estilo para Python (Basada en Google Style Guide)

Este proyecto sigue las normas de codificación de la Guía de Estilo de Python de Google. El objetivo es mantener un código limpio, legible y consistente.

## 1. Reglas de Formato

* **Indentación:** Usar 4 espacios por nivel (no usar tabs).
* **Longitud de línea:** Máximo 80 caracteres.
* **Paréntesis:** Usarlos con moderación. No usarlos en declaraciones `if` o `for` a menos que sea necesario para la continuación de línea.

## 2. Convenciones de Nombres

| Tipo                 | Regla                                    | Ejemplo                    |
| :------------------- | :--------------------------------------- | :------------------------- |
| **Paquetes**   | Minúsculas, sin guiones bajos           | `import mypackage`       |
| **Módulos**   | Minúsculas, pueden llevar guiones bajos | `my_module.py`           |
| **Clases**     | CapWords (PascalCase)                    | `class MyClass:`         |
| **Funciones**  | minúsculas_con_guiones                  | `def calculate_total():` |
| **Variables**  | minúsculas_con_guiones                  | `user_name = "Alex"`     |
| **Constantes** | MAYÚSCULAS_CON_GUIONES                  | `MAX_RETRY = 5`          |

## 3. Comentarios y Docstrings

Todas las funciones y clases deben tener un Docstring siguiendo el formato de Google:

```python
def fetch_data(url: str, retry_count: int = 3) -> dict:
    """Obtiene datos de una API externa.

    Args:
        url: La dirección URL para la petición.
        retry_count: Número de reintentos permitidos.

    Returns:
        Un diccionario con la respuesta procesada.
    """
```

## 4. Importaciones

Las importaciones deben estar en líneas separadas.

Orden:

Librerías estándar (ej. os, sys).

Librerías de terceros (ej. pandas, requests).

Aplicaciones locales del proyecto.

5. Tipado (Type Hinting)
   Se recomienda el uso de anotaciones de tipo para mejorar la claridad y el uso de linters.


This project follows:

- Google Python Style Guide
  https://google.github.io/styleguide/pyguide.html
