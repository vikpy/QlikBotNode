/**
 * @module routes/index
 * @author yianni.ververis@qlik.com
 * @description
 * Handle all of the public front end routes 
*/

const express = require('express');
const router = express.Router();
const path = require('path');

// frontend routes =========================================================
// route to handle all angular requests
router.get('*', function(req, res) {
	res.sendFile(path.join(__dirname, '../../public', 'index.htm')); // load our public/index.html file
//	res.redirect('https://www.google.com');
});

module.exports = router;