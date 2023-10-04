/* 

CorelDraw JS macros that converts all text shapes into curves on
all pages of active document 

Version: 1.0.0
 Author: Anton Syuvaev (h8every1)
   Repo: https://github.com/h8every1/coreldraw-text-to-curves-macros

*/

/*
MIT License

Copyright (c) 2022 Anton Syuvaev

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
*/

class TextToCurvesConvertor {
	constructor() {
		for (let i = 1; i <= host.ActiveDocument.Pages.Count; i++) {
			let page = host.ActiveDocument.Pages.Item(i);

			for (let l = 1; l <= page.Layers.Count; l++) {
				this.convertText(page.Layers.Item(l).Shapes);
			}
			
		}
	}

	convertText(shapes) {
		for (let s = 1; s <= shapes.Count; s++) {
			const pc = shapes.Item(s).PowerClip
			let subShapes = null

			if (shapes.Item(s).Shapes.Count > 0) {
				subShapes = shapes.Item(s).Shapes;
			} else if (pc) {
				subShapes = pc.Shapes;
			}

			if (subShapes) {
				this.convertText(subShapes);
			}
		}

		const texts = shapes.FindShapes(undefined, 6, true);

		for (var t = 1; t <= texts.Count; t++) {
			texts.Item(t).ConvertToCurves();
		}
	}
}

new TextToCurvesConvertor();
