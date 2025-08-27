# Signup Service

A minimal signup service that validates email and password requirements.

## Setup

```bash
npm install
```

## Running the service

```bash
npm start
```

The service will start on port 3000 by default.

## API

### POST /api/signup

Creates a new user signup with validation.

**Request Body:**
```json
{
  "email": "user@example.com",
  "password": "password123"
}
```

**Validation Rules:**
- Email must be non-empty
- Password must be at least 8 characters

**Success Response (201):**
```json
{
  "message": "Signup successful"
}
```

**Error Response (400):**
```json
{
  "error": "Email is required"
}
```

or

```json
{
  "error": "Password must be at least 8 characters"
}
```

## Testing

```bash
npm test
```