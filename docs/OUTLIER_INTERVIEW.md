# Preparación para Outlier e entrevistas técnicas

## Presentación de 30 segundos — español

Desarrollé una aplicación en Python para reemplazar parte de un proceso manual de distribución de mercadería. El programa importa remitos desde Excel, agrupa artículos y talles, controla el stock disponible, permite asignar unidades a distintos destinos y guarda el resultado en una base SQLite. Para el portfolio preparé una versión pública con datos ficticios, separé la lógica de negocio de la interfaz y agregué pruebas automatizadas.

## 30-second introduction — English

I developed a Python desktop application to replace part of a manual merchandise allocation workflow. It imports receipt data from Excel, groups products and sizes, validates available stock, lets users allocate units to different destinations, stores the results in SQLite, and exports reports. For my public portfolio, I created a synthetic-data version, separated the business logic from the user interface, and added automated tests for the critical rules.

## Preguntas probables

### Why did you use SQLite?

For the public demo I wanted a self-contained application that anyone could run without credentials or server configuration. In a multi-user production environment, I would use a centralized database such as PostgreSQL or SQL Server.

### How do you prevent over-allocation?

Before saving, the application aggregates every entered quantity by size and compares that total with the initial stock. If any size exceeds its available quantity, the whole operation is rejected and no partial data is committed.

### What happens to unallocated units?

The application calculates the remainder for each size and automatically records it under the warehouse associated with the selected banner. This keeps the total distributed quantity equal to the original stock.

### Why separate the data layer from Tkinter?

It makes the business logic easier to test, reuse, and eventually expose through an API. The interface only collects input and displays results, while the store class owns validation, transactions, SQL, and status updates.

### What would you improve for production?

I would add authentication, role-based permissions, audit history, centralized persistence, concurrency controls, logging, database migrations, and a web interface built with FastAPI and React.

### What was the hardest part?

The hardest part was translating an informal operational process into explicit rules and edge cases that the application could validate consistently.

## Conceptos para repasar

- DataFrames, `groupby`, validación de tipos y valores nulos.
- Transacciones y claves foráneas en SQL.
- Diferencia entre lógica de presentación y lógica de negocio.
- Manejo de excepciones y mensajes al usuario.
- Pruebas unitarias, casos límite e invariantes.
- Complejidad de agrupar y validar cantidades.

## Public project links

- Author: Joaquín Baigorria
- LinkedIn: https://www.linkedin.com/in/joaquin-baigorria-5224b0a0/
- GitHub: https://github.com/joabaigo10
- Repository: https://github.com/joabaigo10/Distribution-Manager
