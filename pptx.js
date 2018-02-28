var officegen = require ('officegen')
var fs = require('fs');
var pptx = officegen('pptx');

var slide;
var pObj;

pptx.on('finalize', function (written) {
  console.log('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n');

  // clear the temporatory files
});

pptx.on('error', function (err) {
  console.log(err);
});

pptx.setDocTitle('Sample PPTX Document');

const generateExampleSlides = (data, callback) => {
  
  // do the rest things here
  // console.log('finalize');

  // Let's create a new slide:
  slide = pptx.makeNewSlide();

  slide.name = 'The first slide!';

  // Change the background color:
  slide.back = '000000';

  // Declare the default color to use on this slide:
  slide.color = 'ffffff';

  // Basic way to add text string:
  slide.addText('Created using Officegen version ' + officegen.version);
  slide.addText('Fast position', 0, 20);
  slide.addText('Full line', 0, 40, '100%', 20);

  // Add text box with multi colors and fonts:
  slide.addText([
    { text: 'Hello ', options: { font_size: 56 } },
    { text: 'World!', options: { font_size: 56, font_face: 'Arial', color: 'ffff00' } }
  ], { cx: '75%', cy: 66, y: 150 });
  // Please note that you can pass object as the text parameter to addText.

  // For a single text just pass a text string to addText:
  slide.addText('Office generator', { y: 66, x: 'c', cx: '50%', cy: 60, font_size: 48, color: '0000ff' });

  pObj = slide.addText('Boom\nBoom!!!', { y: 100, x: 10, cx: '70%', font_face: 'Wide Latin', font_size: 54, color: 'cc0000', bold: true, underline: true });
  pObj.options.y += 150;

  // 2nd slide:
  slide = pptx.makeNewSlide();

  // For every color property (including the back color property) you can pass object instead of the color string:
  slide.back = { type: 'solid', color: '004400' };
  pObj = slide.addText('Office generator', { y: 'c', x: 0, cx: '100%', cy: 66, font_size: 48, align: 'center', color: { type: 'solid', color: '008800' } });
  pObj.setShadowEffect('outerShadow', { bottom: true, right: true });

  slide = pptx.makeNewSlide();

  slide.show = false;
  slide.addText('Red line', 'ff0000');
  slide.addShape(pptx.shapes.OVAL, { fill: { type: 'solid', color: 'ff0000', alpha: 50 }, line: 'ffff00', y: 50, x: 50 });
  slide.addText('Red box 1', { color: 'ffffff', fill: 'ff0000', line: 'ffff00', line_size: 5, y: 100, rotate: 45 });
  slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 150, cy: 0, cx: 300 });
  slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 150, cy: 100, cx: 0 });
  slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 249, x: 150, cy: 0, cx: 300 });
  slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 449, cy: 100, cx: 0 });
  slide.addShape(pptx.shapes.LINE, { line: '000088', y: 150, x: 150, cy: 100, cx: 300 });
  slide.addShape(pptx.shapes.LINE, { line: '000088', y: 150, x: 150, cy: 100, cx: 300 });
  slide.addShape(pptx.shapes.LINE, { line: '000088', y: 170, x: 150, cy: 100, cx: 300, line_head: 'triangle' });
  slide.addShape(pptx.shapes.LINE, { line: '000088', y: 190, x: 150, cy: 100, cx: 300, line_tail: 'triangle' });
  slide.addShape(pptx.shapes.LINE, { line: '000088', y: 210, x: 150, cy: 100, cx: 300, line_head: 'stealth', line_tail: 'stealth' });
  pObj = slide.addShape(pptx.shapes.LINE);
  pObj.options.line = '008888';
  pObj.options.y = 210;
  pObj.options.x = 150;
  pObj.options.cy = 100;
  pObj.options.cx = 300;
  pObj.options.line_head = 'stealth';
  pObj.options.line_tail = 'stealth';
  pObj.options.flip_vertical = true;
  slide.addText('Red box 2', { color: 'ffffff', fill: 'ff0000', line: 'ffff00', y: 350, x: 200, shape: pptx.shapes.ROUNDED_RECTANGLE, indentLevel: 1 });

  slide = pptx.makeNewSlide();
  data.forEach(commit => {
    slide.addText(commit.node.message)
  })

  callback(pptx);
}

module.exports.generateExampleSlides = generateExampleSlides

// function finalize() {
//   var out = fs.createWriteStream('out.pptx');

//   out.on('error', function (err) {
//     console.log(err);
//   });

//   pptx.generate(out);
// }

// generateExampleSlides(function () {
//   var out = fs.createWriteStream('out.pptx');

//   out.on('error', function (err) {
//     console.log(err);
//   });

//   pptx.generate(out);
// })