const fs = require('fs');
const path = require('path');

const filePath = 'C:\\Users\\USER\\OneDrive\\Anti-Gravity\\고용관련 세액공제\\tax-credit-app\\server\\index.js';
let content = fs.readFileSync(filePath, 'utf8');

// Fix api/refresh scanDirs loop
content = content.replace(/scanDirs\.forEach\(dir => \{([\s\S]+?)\}\);/g, (match, body) => {
    if (body.includes('api/refresh')) return match; // Avoid accidental match if any
    return `for (const dir of scanDirs) {${body}}`;
});

// Second attempt at fixing the loop if the first regex was too broad/narrow
// specifically target the broken part
content = content.replace(/scanDirs\.forEach\(dir => \{\s+if \(fs\.existsSync\(dir\)\) \{([\s\S]+?)\}\s+\}\);/g, (match, body) => {
    return `for (const dir of scanDirs) { if (fs.existsSync(dir)) {${body}} }`;
});

// Since I have nested loops, regex might be tricky. Let's do a more robust string replacement for the specific areas.
// 1. api/refresh
content = content.replace('scanDirs.forEach(dir => {', 'for (const dir of scanDirs) {');
content = content.replace('});\n\n        console.log(`Reset triggered:', '}\n\n        console.log(`Reset triggered:');

// 2. api/delete-branch & api/clear-branch-data
// Let's just fix the forEach to for...of and closures
content = content.replace(/scanDirs\.forEach\(dir => \{([\s\S]+?)movedCount\+\+;\s+\}\s+\}\s+\}\);/g, (match, body) => {
    return `for (const dir of scanDirs) {${body}movedCount++; } } }`;
});

fs.writeFileSync(filePath, content);
console.log('Fixed server/index.js syntax');
