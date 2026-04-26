<?php
header('Content-Type: application/json');
echo json_encode([
    'PATH_INFO' => isset($_SERVER['PATH_INFO']) ? $_SERVER['PATH_INFO'] : 'NOT SET',
    'REQUEST_URI' => $_SERVER['REQUEST_URI'],
    'SCRIPT_NAME' => $_SERVER['SCRIPT_NAME'],
    'PHP_SELF' => $_SERVER['PHP_SELF']
], JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
