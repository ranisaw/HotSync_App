# CustomerApi — Fennex Task 5

A **.NET 10 Web API** providing full CRUD operations over the `dbo.Customers` table in Azure SQL, populated by the Hotsync console app (Task 4).

Built with **raw ADO.NET** (`Microsoft.Data.SqlClient`) — no Entity Framework. All SQL uses parameterised queries.

---

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Access to the Azure SQL `HotsyncDB` database

---

## Configuration

The connection string is stored in `appsettings.json` under the `Database` section and is **never hardcoded** in source code.

```json
{
  "Database": {
    "ConnectionString": "Server=tcp:<your-server>.database.windows.net,1433;Initial Catalog=HotsyncDB;..."
  }
}
```

The `appsettings.json` already contains the correct Azure SQL connection string for the Fennex HotsyncDB.

---

## Running the API

```bash
cd "D:\React projects\New folder"
dotnet run
```

The API will start on:
- **HTTP:**  `http://localhost:5000`
- **HTTPS:** `https://localhost:5001`

---

## Web UI Dashboard

A dedicated, responsive **Fennex Customer Hub** single-page application is hosted directly by the ASP.NET Core web server.

Open your browser and navigate to:
```
http://localhost:5000/
```
(or `https://localhost:5001/` when running with HTTPS)

### Features included in the UI:
- **Live Fleet KPI Metrics**: Real-time summary cards for Total Customers, Active Oil Rigs, Jack-Up Units, and MODUs.
- **Interactive Data Views**: Toggle between a structured **Table View** and an offshore **Grid Cards View**.
- **Real-Time Search & Filtering**: Filter by fleet assets (Rigs, Jack-Ups, MODUs) or search by name, SharePoint ID, address, and URLs.
- **Sorting**: Sort by SQL Id, Company Name, Fleet Size, or Recent Sync timestamp.
- **Full CRUD Modals**:
  - **Create Customer**: Modal with client-side validation rules matching backend constraints.
  - **Edit Customer**: Modify existing customer attributes via `PUT /api/customers/{id}`.
  - **Details Inspector**: Detailed view of all 11 SQL columns, SharePoint metadata, and external site links.
  - **Permanent Deletion**: Protected modal to safely execute `DELETE /api/customers/{id}`.
- **Resilience / Demo Sandbox Mode**: Built-in preview toggle to explore and test the dashboard even if Azure SQL network firewall blocks current IP.

---

## Swagger UI

Open your browser and navigate to:

```
http://localhost:5000/swagger
```

All 5 endpoints are listed and fully testable from the Swagger UI.

---

## Endpoints

### `GET /api/customers`
Returns all customers ordered by `Id` ascending.

**Response: 200 OK**
```json
[
  {
    "id": 1,
    "sharePointId": 42,
    "name": "Acme Drilling Co.",
    "address": "123 Oil Street, Houston TX",
    "numberOfRigs": 5,
    "numberOfJackUps": 2,
    "numberOfModus": 1,
    "siteUrls": "https://acme.sharepoint.com/sites/ops",
    "spLastModified": "2024-06-01T10:00:00Z",
    "syncedAt": "2024-06-15T08:30:00Z",
    "createdAt": "2024-06-15T08:30:00Z"
  }
]
```

---

### `GET /api/customers/{id}`
Returns a single customer by their SQL primary key.

**Response: 200 OK** — customer JSON (same shape as above)  
**Response: 404 Not Found**
```json
{ "error": "Customer with Id 99 was not found." }
```

---

### `POST /api/customers`
Creates a new customer record.

**Request body:**
```json
{
  "sharePointId": 101,
  "name": "Gulf Rigs Ltd.",
  "address": "456 Platform Ave, Aberdeen",
  "numberOfRigs": 3,
  "numberOfJackUps": 1,
  "numberOfModus": 0,
  "siteUrls": "https://gulfrigs.sharepoint.com/sites/main",
  "spLastModified": "2024-07-01T09:00:00Z"
}
```

| Field | Required | Notes |
|-------|----------|-------|
| `sharePointId` | ✅ Yes | Must be unique |
| `name` | ✅ Yes | Max 255 chars |
| `address` | No | Max 1000 chars |
| `numberOfRigs` | No | Must be ≥ 0 |
| `numberOfJackUps` | No | Must be ≥ 0 |
| `numberOfModus` | No | Must be ≥ 0 |
| `siteUrls` | No | Comma-separated |
| `spLastModified` | No | ISO 8601 datetime |

**Response: 201 Created** — full customer JSON with `Location` header pointing to the new resource.  
**Response: 400 Bad Request** — validation errors  
**Response: 409 Conflict** — duplicate `SharePointId`

---

### `PUT /api/customers/{id}`
Updates an existing customer.

**Request body:**
```json
{
  "name": "Gulf Rigs Ltd. (Updated)",
  "address": "789 New Harbour, Aberdeen",
  "numberOfRigs": 4,
  "numberOfJackUps": 2,
  "numberOfModus": 1,
  "siteUrls": "https://gulfrigs.sharepoint.com/sites/main",
  "spLastModified": "2024-08-01T10:00:00Z"
}
```

**Response: 200 OK** — updated customer JSON  
**Response: 400 Bad Request** — validation errors  
**Response: 404 Not Found** — Id does not exist

---

### `DELETE /api/customers/{id}`
Deletes a customer by SQL Id.

**Response: 204 No Content** — deleted successfully  
**Response: 404 Not Found** — Id does not exist

---

### `GET /api/customers/export/excel` (Task 7)
Exports all customer records in a styled tabular layout into an **Excel workbook (.xlsx)** using official **DocumentFormat.OpenXml**.

- **Content-Type**: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- **Filename**: `Fennex_Customers_Export_{timestamp}.xlsx`
- **Columns**: SQL Id, SharePoint ID, Company Name, Address, Oil Rigs, Jack-Up Units, MODU Vessels, Total Fleet, Site URLs, SP Last Modified, Synced At, Created At.
- **Styling**: Bold navy header with white font, zebra alternating rows, customized column widths.

---

### `GET /api/customers/export/word` (Task 8)
Exports consolidated enterprise customer data into a **Word document (.docx)** using **DocumentFormat.OpenXml**.

- **Content-Type**: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- **Filename**: `Fennex_Fleet_Report_{timestamp}.docx`
- **Sections**:
  1. Title Banner & generation metadata
  2. Executive Fleet Summary (total fleets, rigs, Jack-Ups, MODUs)
  3. Consolidated Customer Fleet Table
  4. Detailed Customer Profiles for every enterprise
  5. Confidentiality & security sign-off notice

---

### `GET /api/customers/{id}/export/word` (Task 8 Extension)
Exports an individual customer's executive profile into a formatted **Word document (.docx)** using **DocumentFormat.OpenXml**.

- **Content-Type**: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- **Filename**: `Fennex_Customer_{name}_{id}.docx`

---

## HTTP Status Code Reference

| Code | Meaning |
|------|---------|
| `200 OK` | Successful GET or PUT |
| `201 Created` | Successful POST — includes `Location` header |
| `204 No Content` | Successful DELETE |
| `400 Bad Request` | Validation failed — response includes field errors |
| `404 Not Found` | Resource with given Id does not exist |
| `409 Conflict` | Duplicate `SharePointId` on create |
| `500 Internal Server Error` | Unexpected database or server error |

---

## Project Structure

```
CustomerApi/
├── Controllers/
│   └── CustomersController.cs    # 5 CRUD endpoints — thin, delegates to repository
├── Models/
│   ├── Customer.cs               # Response DTO (all 11 SQL columns)
│   ├── CreateCustomerRequest.cs  # POST body with DataAnnotations validation
│   └── UpdateCustomerRequest.cs  # PUT body with DataAnnotations validation
├── Services/
│   └── CustomerRepository.cs     # All SQL queries via ADO.NET — no EF
├── Middleware/
│   └── ErrorHandlingMiddleware.cs # Global exception handler
├── appsettings.json              # Connection string config
├── CustomerApi.csproj
└── README.md
```

---

## NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| `Microsoft.Data.SqlClient` | 7.1.1 | Azure SQL database access |
| `Swashbuckle.AspNetCore` | 6.9.0 | Swagger / OpenAPI UI |

---

## Security Notes

- All SQL queries use **parameterised `SqlCommand`** — no string concatenation, no SQL injection risk.
- Raw exception messages and stack traces are **never** returned to callers — the global middleware logs full details server-side and returns only a safe JSON error + `traceId`.
- Connection string is read from `appsettings.json` — **never hardcoded**.
