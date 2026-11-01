# Third-Party Software Notices and Licenses

This repository and the FastMailSearch software utilize third-party open-source libraries and components. Each component is acknowledged below alongside its respective copyright notice, license type, and terms.

---

## Summary of Third-Party Dependencies

| Component | Version | License | Usage Scope |
| :--- | :--- | :--- | :--- |
| **Newtonsoft.Json** | 13.0.3 | MIT License | Runtime JSON Serialization (`FastMailSearch.Core`, `UI`) |
| **SQLitePCLRaw** | 2.1.6 | Apache License 2.0 | Runtime SQLite Native Provider Abstraction (`Infrastructure`, `UI`) |
| **SQLite (e_sqlite3)** | 3.x | Public Domain / SQLite Blessing | Embedded Native Database Engine |
| **System.Data.SQLite.Core** | 1.0.118 | Public Domain / Dedicated | COM Add-In SQLite Integration |
| **Microsoft.Data.Sqlite** | 8.0.3 | MIT License | Database Integration & Performance Benchmarks |
| **System.Management** | 8.0.0+ | MIT License | Windows Hardware ID Resolution (`Infrastructure`) |
| **System.Security.Cryptography.ProtectedData** | 8.0.0+ | MIT License | Local DPAPI Machine Cryptography (`Infrastructure`) |
| **Microsoft.Office.Interop.Outlook** | 15.0.4797.1004 | Microsoft Software License | Office/Outlook COM Interop Metadata |
| **xUnit.net** | 2.7.0 | Apache License 2.0 | Automated Test Framework (`tests/`) |
| **Microsoft.NET.Test.Sdk** | 17.9.0 | MIT License | Test Execution Harness (`tests/`) |

---

## License Texts and Attributions

### 1. Newtonsoft.Json
- **Copyright:** (c) 2007 James Newton-King
- **License:** MIT License
- **URL:** https://github.com/JamesNK/Newtonsoft.Json

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

---

### 2. SQLitePCLRaw
- **Copyright:** (c) 2014-2023 Eric Sink
- **License:** Apache License, Version 2.0
- **URL:** https://github.com/ericsink/SQLitePCL.raw

Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. You may obtain a copy of the License at:
http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.

---

### 3. SQLite and System.Data.SQLite
- **Authors:** D. Richard Hipp and the SQLite Development Team
- **License:** Public Domain
- **Notice:** The SQLite software is in the public domain, dedicated by its authors to everyone around the world. No license fee is required.
  > "May you do good and not evil.  
  > May you find forgiveness for yourself and forgive others.  
  > May you share freely, never taking more than you give."

---

### 4. Microsoft .NET Libraries (Microsoft.Data.Sqlite, System.Management, System.Security.Cryptography.ProtectedData, Microsoft.NET.Test.Sdk)
- **Copyright:** (c) .NET Foundation and Contributors.
- **License:** MIT License
- **URL:** https://github.com/dotnet/runtime

Licensed under the MIT License. Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files, to deal in the Software without restriction, subject to the standard MIT conditions.

---

### 5. xUnit.net
- **Copyright:** (c) 2014-2024 .NET Foundation and Contributors.
- **License:** Apache License, Version 2.0
- **URL:** https://github.com/xunit/xunit

Licensed under the Apache License, Version 2.0 (the "License"). You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0.

