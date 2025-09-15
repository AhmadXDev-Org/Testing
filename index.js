const express = require('express');
const app = express();

// Middleware to parse JSON bodies
app.use(express.json());

// POST /api/signup endpoint
app.post('/api/signup', (req, res) => {
    const { email, password } = req.body;
    
    // Validate email is non-empty
    if (!email || email.trim() === '') {
        return res.status(400).json({ error: "Email is required" });
    }
    
    // Validate password is at least 8 characters
    if (!password || password.length < 8) {
        return res.status(400).json({ error: "Password must be at least 8 characters" });
    }
    
    // If validation passes, return success
    res.status(201).json({ message: "Signup successful" });
});

// Start server only if this file is run directly (not in tests)
if (require.main === module) {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
        console.log(`Server running on port ${PORT}`);
    });
}

module.exports = app;