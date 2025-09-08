License Information for AP Excel Agent POC

This project is distributed under a dual-license model. Different parts of the codebase are subject to different licenses due to the licenses of its dependencies. Please read the following sections carefully.

1. Frontend Application License (GNU GPLv3)

The frontend code, which includes the index.html file and all JavaScript code running in the browser, is a derivative work of the HyperFormula library. Due to the "copyleft" nature of the GNU General Public License v3, the frontend application is also licensed under the GNU GPLv3.

What this means: If you create and distribute a modified version of this frontend application, your modified version must also be licensed under the GNU GPLv3 and you must make the source code available.

A full copy of the GNU GPLv3 is included below. You can also view it online at https://github.com/handsontable/hyperformula/blob/master/LICENSE.txt.

<details>
<summary>Click to view the full text of the GNU GPLv3</summary>

Copyright (c) HANDSONCODE sp. z o. o.

HYPERFORMULA is a software distributed by HANDSONCODE sp. z o. o., a
Polish corporation based in Gdynia, Poland, at Aleja Zwyciestwa 96-98,
registered by the District Court in Gdansk under number 538651,
EU VAT: PL5862294002, share capital: PLN 62,800.00.

This software is dual-licensed, giving you the option to use it under
either a proprietary license or the GNU General Public License version 3
(GPLv3). The specific license under which you use the software is
determined by the license key you apply. Each licensing option comes with
its own terms and conditions as specified below.

  1. PROPRIETARY LICENSE:

    Your use of this software is subject to the terms included in an
    applicable proprietary license agreement between you and HANDSONCODE.
    The proprietary license can be purchased from HANDSONCODE or an
    authorized reseller.

  2. GNU GENERAL PUBLIC LICENSE v3:

    This software is also available under the terms of the GNU General
    Public License v3. You are permitted to run, modify, and distribute
    this software under the terms of the GPLv3, as published by the Free
    Software Foundation. The full text of the GPLv3 can be found at
    https://www.gnu.org/licenses/.

UNLESS EXPRESSLY AGREED OTHERWISE, HANDSONCODE PROVIDES THIS SOFTWARE ON
AN "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, AND IN NO
EVENT AND UNDER NO LEGAL THEORY, SHALL HANDSONCODE BE LIABLE TO YOU FOR
DAMAGES, INCLUDING ANY DIRECT, INDIRECT, SPECIAL, INCIDENTAL, OR
CONSEQUENTIAL DAMAGES OF ANY CHARACTER ARISING FROM USE OR INABILITY TO
USE THIS SOFTWARE.
</details>

2. Backend Server License (MIT)

All original backend server code, located in the src/ directory (including index.ts, agent.ts, sheet.ts, etc.), is licensed under the permissive MIT License. This code is separate from the frontend and does not directly link to any GPL-licensed libraries.

Copyright (c) 2025 Octalatiq LLC

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

3. Third-Party Dependencies & Usage

This project relies on third-party components with their own licenses that you must comply with.

    Handsontable: The spreadsheet grid UI is used under a Non-commercial license. Any use of this project in a commercial application requires a commercial license from Handsontable. More info: handsontable.com/pricing

    HyperFormula: The formula engine is used under the GNU GPLv3. Use in any project that is not also licensed under the GPLv3 requires a commercial license. More info: handsontable.com/hyperformula

    Ollama: The local LLM runner is open-source and distributed under the MIT License. More info: github.com/ollama/ollama

    Google Gemini API: Use of the Gemini API is governed by the Google AI Studio and Gemini API Terms of Service. You must comply with these terms, including any usage policies and restrictions. More info: ai.google.dev/terms

    Qwen Models (e.g., qwen3:32b): The Qwen models developed by Alibaba Cloud are subject to the APACHE 2.0 License agreement. This license has specific terms, including conditions for commercial use. You must review and comply with this license. More info: https://huggingface.co/Qwen/Qwen3-32B/blob/main/LICENSE
    
    PDF.js (`pdfjs-dist`): The PDF parsing library is provided by Mozilla under the **Apache License 2.0**. This is a permissive license allowing commercial use. More info: [github.com/mozilla/pdf.js](https://github.com/mozilla/pdf.js)
    
    SheetJS (`xlsx`): The XLSX file parser and writer is provided under the **Apache License 2.0**. This is a permissive license allowing commercial use. More info: [sheetjs.com](https://sheetjs.com/)

By using, modifying, or distributing this project, you agree to be bound by the terms of all applicable licenses mentioned above.