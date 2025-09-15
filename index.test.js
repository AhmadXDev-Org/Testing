const request = require('supertest');
const app = require('./index');

describe('POST /api/signup', () => {
    test('should return 400 error for empty email', async () => {
        const response = await request(app)
            .post('/api/signup')
            .send({ email: '', password: 'password123' })
            .expect(400);
        
        expect(response.body).toEqual({ error: "Email is required" });
    });

    test('should return 400 error for missing email', async () => {
        const response = await request(app)
            .post('/api/signup')
            .send({ password: 'password123' })
            .expect(400);
        
        expect(response.body).toEqual({ error: "Email is required" });
    });

    test('should return 400 error for email with only whitespace', async () => {
        const response = await request(app)
            .post('/api/signup')
            .send({ email: '   ', password: 'password123' })
            .expect(400);
        
        expect(response.body).toEqual({ error: "Email is required" });
    });

    test('should return 400 error for short password', async () => {
        const response = await request(app)
            .post('/api/signup')
            .send({ email: 'test@example.com', password: '1234567' })  // 7 chars
            .expect(400);
        
        expect(response.body).toEqual({ error: "Password must be at least 8 characters" });
    });

    test('should return 400 error for missing password', async () => {
        const response = await request(app)
            .post('/api/signup')
            .send({ email: 'test@example.com' })
            .expect(400);
        
        expect(response.body).toEqual({ error: "Password must be at least 8 characters" });
    });

    test('should return 201 for valid email and password', async () => {
        const response = await request(app)
            .post('/api/signup')
            .send({ email: 'test@example.com', password: 'password123' })
            .expect(201);
        
        expect(response.body).toEqual({ message: "Signup successful" });
    });
});