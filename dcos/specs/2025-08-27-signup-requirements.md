# Signup Requirements (MVP)
- Endpoint: POST /api/signup
- Validate: email must be non-empty; password >= 8 chars
- Error format: { "error": "<message>" }
- Tests: unit tests for empty email + short password
