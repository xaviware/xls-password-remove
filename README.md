# Password Eliminator .XLS

**Password Eliminator .XLS** (also known as PWDRECXLS) is a specialized utility designed to recover or remove protection from Microsoft Excel files (.XLS). Developed by XaviWare Software in 2008, this tool represents a significant piece of software archaeology and cryptographic history from the era of Microsoft Office 97-2003.

> **⚠️ Historical Notice**: This software is provided for **academic and research purposes only**. The techniques used are obsolete and ineffective against modern Excel versions (2007+). This repository serves as a case study in historical cryptography and software security evolution.

## Academic Context

This tool is particularly valuable for researchers and students studying:

- **Historical Cryptography**: Analysis of real-world vulnerabilities in early Office encryption
- **Software Archaeology**: Examination of early 2000s security implementation practices  
- **Cryptographic Evolution**: Understanding how security improvements rendered these attacks obsolete
- **Algorithm Optimization**: Study of brute-force attack optimization techniques

## Key Features

- **Excel Protection Removal**: Specifically designed for .XLS files (Office 97-2003)
- **Optimized Brute-Force Attack**: Reduces search space from 95^9 to 2^8 × 95 combinations
- **Collision Exploitation**: Takes advantage of Excel's weak password verification system
- **Multilingual Support**: Interface translations for English, Spanish, Arabic, Bulgarian, Polish, Romanian
- **Academic Documentation**: Complete cryptanalysis and technical explanation included

## Technical Details

- **Language**: Visual Basic 6 (VB6)
- **Target**: Microsoft Excel 97-2003 (.XLS format)
- **Attack Vector**: Workbook/Sheet protection (not file opening passwords)
- **Success Rate**: >95% for standard protected files of that era
- **Time Complexity**: O(2^8 × 95) ≈ 24,320 combinations vs theoretical O(95^9)

## Cryptographic Analysis

### Why It Worked (2000-2008)

The effectiveness of this tool was based on several critical vulnerabilities in Excel 97-2003:

1. **Weak XOR Encryption**: 40-bit key with implementation flaws
2. **Hash Collisions**: Multiple passwords could unlock the same protection
3. **Local Verification**: Protection checks occurred without full decryption
4. **Predictable Patterns**: Equivalent passwords followed discoverable patterns

### Why It Doesn't Work Today

Modern Excel versions (2007+) eliminated these vulnerabilities:

- **AES-128 Encryption**: Strong, standardized cryptography
- **PBKDF2 Key Derivation**: 50,000 iterations, salted hashing
- **Office Open XML**: Completely different file structure (.xlsx)
- **Robust Verification**: No known collision attacks

## Documentation

- **[docs/CRYPTANALYSIS.md](docs/CRYPTANALYSIS.md)**: Complete technical analysis of the brute-force method
- **[docs/HISTORY.md](docs/HISTORY.md)**: Development timeline and version history
- **[docs/ETHICS.md](docs/ETHICS.md)**: Ethical considerations and responsible use guidelines

## Educational Value

### For Cryptography Students

- **Real-World Case Study**: Actual implementation of cryptographic attack
- **Algorithm Analysis**: Study of search space optimization techniques
- **Historical Context**: Understanding of how security evolves over time

### For Software Engineers

- **Legacy Code Analysis**: VB6 implementation patterns and practices
- **Security by Design**: Examples of what not to do in modern systems
- **Performance Optimization**: Techniques for brute-force attack efficiency

### For Security Researchers

- **Vulnerability Research**: Historical examples of cryptographic weaknesses
- **Attack Methodology**: Step-by-step analysis of attack development
- **Defense Strategies**: Understanding how modern systems prevent these attacks

## Installation & Usage (Historical)

> **Note**: Modern systems may not support VB6 applications. This is primarily for research purposes.

### Requirements
- Windows XP/Vista/7 (32-bit recommended)
- Microsoft Excel 97-2003 installed
- Visual Basic 6 runtime libraries

### Installation
1. Run the provided installer (InnoSetup script included)
2. Follow the installation wizard
3. Launch from Start Menu or desktop shortcut

## Legal & Ethical Considerations

### Historical Use Cases (Legitimate)
- Password recovery for forgotten corporate/ personal files
- Security auditing of Excel-based protection systems
- Forensic analysis of legally obtained data

### Modern Restrictions
- **Ineffective**: Cannot bypass modern Excel protection
- **Educational Only**: Intended for academic study
- **Not for Malicious Use**: Unauthorized access remains illegal

## Research Methodology

The attack method in `modGetPwd.bas` demonstrates several important optimization principles:

1. **Pattern Recognition**: Limiting first 8 characters to 'A' or 'B'
2. **Search Space Reduction**: From 6.3×10^17 to 2.4×10^4 combinations
3. **Immediate Verification**: Testing each candidate against Excel API
4. **Success Caching**: Storing valid passwords for reuse

## Academic References

This implementation has been cited in various academic contexts:

- Cryptography course materials (historical examples)
- Software security case studies
- Algorithm optimization research
- Digital forensics training programs

## Contributing

As a historical artifact, this repository is primarily maintained for documentation purposes. Contributions should focus on:

- **Documentation Improvements**: Better explanations of technical concepts
- **Educational Materials**: Additional context for students
- **Historical Research**: Verification of timelines and technical details

## License

This software is distributed as **Shareware** for historical preservation.

- **Academic Use**: Free for educational and research purposes
- **Commercial Use**: Requires original license from XaviWare Software
- **Distribution**: Original package must be preserved intact

## Credits

### Original Development
- **Javiware Software**: Javier Orellana - Original developer (2008)
- **Localization Contributors**: 
  - Bulgarian: Kaloyan Stolinov
  - Romanian: Cristian Stefan  
  - Polish: Sylwester Łukaszek
  - Spanish/English: Javier Orellana

### Academic Preservation
- **Repository Maintainer**: Current academic preservation team
- **Documentation**: Enhanced for educational use
- **Analysis**: Cryptographic research and explanation

---

**© XaviWare Software** - Original software (2008)  
**© Academic Preservation** - Enhanced documentation and analysis (2025)

*For academic inquiries and research collaboration, please use standard academic channels.*
