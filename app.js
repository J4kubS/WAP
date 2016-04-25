var J4_Spreadsheet = (function() {

/*
 * CONSTANTS
 *
*/

var ASCII_A = "A".charCodeAt(0),
	KEY_BACKSPACE =  8,
	KEY_ENTER     = 13,
	KEY_ESCAPE    = 27,
	KEY_LEFT      = 37,
	KEY_UP        = 38,
	KEY_RIGHT     = 39,
	KEY_DOWN      = 40,
	KEY_DELETE    = 46,
	Config = {
		patterns: {
			avg: /AVG\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/g,
			sum: /SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/g,
			ref: /\b([A-Z]+)(\d+)\b/g
		},
		css: {
			container    : "j4spreadsheet-container",
			spreadsheet  : "j4spreadsheet-spreadsheet",
			emptyHeader  : "j4spreadsheet-header-empty",
			columnHeader : "j4spreadsheet-header-column",
			rowHeader    : "j4spreadsheet-header-row",
			header       : "j4spreadsheet-header",
			cellSelected : "j4spreadsheet-cell-selected",
			cellEditing  : "j4spreadsheet-cell-editing",
			cellValue    : "j4spreadsheet-cell-value",
			cell         : "j4spreadsheet-cell"
		}
	};

/*
 * END CONSTANTS
 *
*/

/*
 * HELPERS
 *
*/

// Source: toddmotto.com/ditch-the-array-foreach-call-nodelist-hack/
function forEach(array, callback, scope) {
	var i;

	for (i = 0; i < array.length; i += 1) {
		callback.call(scope, i, array[i]);
	}
}

// Source: stackoverflow.com/questions/1303646/check-whether-variable-is-number-or-string-in-javascript
function isNumber(o) {
	return ! isNaN (o - 0) && o !== "" && o !== false;
}

function isArrowKey(key) {
	return key == KEY_LEFT || key == KEY_UP || key == KEY_RIGHT || key == KEY_DOWN;
}

function getAncestorByTagName(node, tagName) {
	if (node && tagName) {
		tagName = tagName.toLowerCase();

		while (node && node.tagName.toLowerCase() !== tagName) {
			node = node.parentNode;
		}

		return node;
	}
}

function getColumnIndex(node) {
	node = getAncestorByTagName(node, "td");

	if (node) {
		return node.cellIndex;
	}
}

function getRowIndex(node) {
	node = getAncestorByTagName(node, "tr");

	if (node) {
		return node.rowIndex;
	}
}

// Encode 0 based column index as [A-Z]+ string.
function encodeColumn(column) {
	var div = Math.floor(column / 26),
		rem = column % 26;

	if (div > 0) {
		return encodeColumn(div - 1) + String.fromCharCode(ASCII_A + rem);
	} else {
		return String.fromCharCode(ASCII_A + rem);
	}
}

// Decode 0 based column index from [A-Z]+ string.
function decodeColumn(column) {
	var result = 0,
		i;

	for (i = 0; i < column.length; i += 1) {
		result = 26 * result + (column.charCodeAt(i) - ASCII_A + 1);
	}

	return result - 1;
}

/*
 * END HELPERS
 *
*/

/*
 * OBSERVABLE
 *
*/

var Observable = function() {
	this._observers = [];
};

Observable.prototype.observe = function(observer) {
	if (typeof observer === "function" && this._observers.indexOf(observer) < 0) {
		this._observers.push(observer);
	}
};

// Notify all registered observers. Any arguments are passed to them.
Observable.prototype.notify = function() {
	var args = Array.prototype.slice.call(arguments, 0);

	this._observers.forEach(function(observer) {
		observer.apply(null, args);
	});
};

/*
 * END OBSERVABLE
 *
*/

/*
 * CELL
 *
*/

var Cell = function(row, column) {
	Observable.call(this);

	this._colum = column;
	this._row = row;

	this._selected = false;
	this._editing = false;

	// Other cells that depend on this cell's value.
	this._related = [];
	this._expression = null;
	this._value = null;

	// Reference to a view associated with this cell.
	this._node = null;
};

Cell.prototype = Object.create(Observable.prototype);
Cell.prototype.constructor = Cell;

Cell.prototype.column = function() {
	return this._colum;
};

Cell.prototype.row = function() {
	return this._row;
};

Cell.prototype.selected = function(arg) {
	if (typeof arg === "undefined") {
		return this._selected;
	} else {
		this._selected = arg;
		this.notify({
			type: "selected",
			cell: this
		});

		return this;
	}
};

Cell.prototype.editing = function(arg) {
	if (typeof arg === "undefined") {
		return this._editing;
	} else {
		this._editing = arg;
		this.notify({
			type: "editing",
			cell: this
		});

		return this;
	}
};

Cell.prototype.related = function(arg) {
	if (typeof arg === "undefined") {
		return this._related;
	} else if (this !== arg && this._related.indexOf(arg) == -1) {
		this._related.push(arg);
		return this;
	}
};

Cell.prototype.expression = function(arg) {
	if (typeof arg === "undefined") {
		return this._expression;
	} else {
		this._expression = arg;
		this.notify({
			type: "expression",
			cell: this
		});

		return this;
	}
};

Cell.prototype.value = function(arg) {
	if (typeof arg === "undefined") {
		return this._value;
	} else {
		this._value = arg;
		this.notify({
			type: "value",
			cell: this
		});

		return this;
	}
};

Cell.prototype.node = function(arg) {
	if (typeof arg === "undefined") {
		return this._node;
	} else {
		this._node = arg;
		return this;
	}
};

/*
 * END CELL
 *
*/

/*
 * SPREADSHEET
 *
*/

var Spreadsheet = function(numRows, numColumns) {
	var cell,
		i, j;

	Observable.call(this);

	this._numColumns = numColumns;
	this._numRows = numRows;
	this._cells = [];

	// Evaluate cell (or depended cells) if needed and pass the event on.
	var observer = function(event) {
		if (event.type === "expression" && event.cell.expression()) {
			evaluate(event.cell);
		} else if (event.type === "value") {
			event.cell.related().forEach(evaluate);
		}

		this.notify(event);
	}.bind(this);

	// Evaluate cell's expression.
	//
	// WARNING: Expression is evaluated using eval!
	var evaluate = function(cell) {
		var expression = cell.expression() && cell.expression().substring(1) || null;

		if (! expression) {
			cell.value(null);
			return;
		}

		expression = evaluateAvg(cell, expression);
		expression = evaluateSum(cell, expression);
		expression = evaluateRef(cell, expression);

		try {
			// Yup. It's here.
			cell.value(eval(expression));
		} catch (e) {
			if (e instanceof SyntaxError || e instanceof ReferenceError) {
				cell.value("#EXPRESSION");
			} else {
				throw e;
			}
		}
	}.bind(this);

	// Replace any AVG() calls with it's actual value.
	var evaluateAvg = function(cell, expression) {
		return expression.replace(Config.patterns.avg, function(match, col1, row1, col2, row2) {
			var cell1 = this.cell(row1 - 1, decodeColumn(col1)),
				cell2 = this.cell(row2 - 1, decodeColumn(col2)),
				sum = 0, range, i;

			// Check if the range is valid and not self-referencing.
			if (cell1 && cell2 && cell1 !== cell && cell2 !== cell) {
				range = this._range(cell1, cell2);

				for (i = 0; i < range.length; i += 1) {
					sum += parseInt(range[i].value());
					range[i].related(cell);
				}

				match = sum / range.length;
			}

			return match;
		}.bind(this));
	}.bind(this);

	// Replace any SUM() calls with it's actual value.
	var evaluateSum = function(cell, expression) {
		return expression.replace(Config.patterns.sum, function(match, col1, row1, col2, row2) {
			var cell1 = this.cell(row1 - 1, decodeColumn(col1)),
				cell2 = this.cell(row2 - 1, decodeColumn(col2)),
				range, i;

			if (cell1 && cell2 && cell1 !== cell && cell2 !== cell) {
				range = this._range(cell1, cell2);
				match = 0;

				for (i = 0; i < range.length; i += 1) {
					match += parseInt(range[i].value());
					range[i].related(cell);
				}
			}

			return match;
		}.bind(this));
	}.bind(this);

	// Replace any cell references with their actual value.
	var evaluateRef = function(cell, expression) {
		return expression.replace(Config.patterns.ref, function(match, col, row) {
			var targetCell = this.cell(row - 1, decodeColumn(col));

			if (targetCell && targetCell !== cell) {
				match = targetCell.value();
				targetCell.related(cell);

				if (! isNumber(match)) {
					// Add quotes to string values and escape any quotes in them.
					match = "'" + match.replace(/'/g, "\\'").replace(/"/g, '\\"') + "'";
				}
			}

			return match;
		}.bind(this));
	}.bind(this);

	for (i = 0; i < numRows; i += 1) {
		for (j = 0; j < numColumns; j += 1) {
			cell = new Cell(i, j);

			cell.observe(observer);
			this._cells.push(cell);
		}
	}
};

Spreadsheet.prototype = Object.create(Observable.prototype);
Spreadsheet.prototype.constructor = Spreadsheet;

Spreadsheet.prototype.cell = function(row, column) {
	var index;

	if (row >= 0 && row < this._numRows && column >= 0 && column < this._numColumns) {
		index = row * this._numColumns + column;

		if (index < this._cells.length) {
			return this._cells[index];
		}
	}
};

Spreadsheet.prototype.selectionExpression = function(expression) {
	this._cells.forEach(function(cell) {
		if (cell.selected()) {
			cell.expression(expression);
		}
	});

	return this;
};

Spreadsheet.prototype.selectionValue = function(value) {
	this._cells.forEach(function(cell) {
		if (cell.selected()) {
			cell.value(value);
		}
	});

	return this;
};

Spreadsheet.prototype.selectionClear = function() {
	this._cells.forEach(function(cell) {
		if (cell.selected()) {
			cell.selected(false);
		}
	});

	return this;
};

// If called with one argument, clear current selection and set that cell as selected.
// If called with two arguments, clear current selection and set all cells in
// rectangle, where two provided cells serve as corners, as selected.
Spreadsheet.prototype.select = function(cell1, cell2) {
	var range;

	if (typeof cell2 === "undefined") {
		this.selectionClear();
		cell1.selected(true);
	} else {
		range = this._range(cell1, cell2);

		this._cells.forEach(function(cell) {
			if (range.indexOf(cell) == -1) {
				cell.selected(false);
			} else {
				cell.selected(true);
			}
		});
	}

	return this;
};

// Get all cells inside the rectangle where cell1 are cell2 corners.
Spreadsheet.prototype._range = function(cell1, cell2) {
	var columnFrom, columnTo,
		rowFrom, rowTo,
		range = [],
		i, j;

	columnFrom = Math.min(cell1.column(), cell2.column()),
	columnTo = Math.max(cell1.column(), cell2.column()),
	rowFrom = Math.min(cell1.row(), cell2.row()),
	rowTo = Math.max(cell1.row(), cell2.row());

	for (i = rowFrom; i <= rowTo; i += 1) {
		for (j = columnFrom; j <= columnTo; j += 1) {
			range.push(this.cell(i, j));
		}
	}

	return range;
};

/*
 * END SPREADSHEET
 *
*/

/*
 * CONTROLLER
 *
*/

var Controller = (function() {

	// Attach new controller the view and the model.
	function init(view, model) {
		var scope = {
			isMouseDown: false,
			cellFrom: null,
			cellTo: null,

			model: model,
			view: view
		};

		model.observe(modelObserver.bind(scope));

		view.addEventListener("unload", done.bind(scope));
		view.addEventListener("mousedown", onMouseDown.bind(scope));
		view.addEventListener("mouseup", onMouseUp.bind(scope));
		view.addEventListener("mouseleave", onMouseLeave.bind(scope));
		view.addEventListener("mousemove", onMouseMove.bind(scope));
		view.addEventListener("dblclick", onClick.bind(scope));
		view.addEventListener("click", onClick.bind(scope));
		view.addEventListener("keypress", onKeyPress.bind(scope));
		view.addEventListener("keydown", onKeyDown.bind(scope));
		view.addEventListener("blur", onLostFocus.bind(scope), true);
		view.addEventListener("focusout", onLostFocus.bind(scope));

		forEach(view.getElementsByClassName(Config.css.cellValue), function(index, node) {
			var column = getColumnIndex(node) - 1,
				row = getRowIndex(node) - 1;

			// Store a reference to cell's model in cell's view and vice versa.
			// Null both of the references on unload to prevent memory leaks.
			node.J4_cell = model.cell(row, column);
			node.J4_cell.node(node);
		});
	}

	// Null the references between cell's view and model.
	function done() {
		forEach(this.view.getElementsByClassName(Config.css.cellValue), function(index, node) {
			if (node.J4_cell) {
				node.J4_cell.node(null);
				node.J4_cell = null;
			}
		});
	}

	/* HANDLERS */

	function modelObserver(event) {
		var cell = event.cell,
			node = cell.node();

		switch (event.type) {
			case "selected":
				if (! cell.selected()) {
					node.classList.remove(Config.css.cellSelected);
				} else {
					node.classList.add(Config.css.cellSelected);
				}
				break;

			case "editing":
				if (! cell.editing()) {
					node.classList.remove(Config.css.cellEditing);
					node.setAttribute("readonly", "readonly");
				} else {
					node.classList.add(Config.css.cellEditing);
					node.removeAttribute("readonly");
					// Because IE is a piece of crap!
					node.focus();
					node.focus();
				}
				break;

			case "expression":
			case "value":
				update(node);
				break;
		}
	}

	function update(target) {
		var value;

		if (target instanceof Cell && target.node()) {
			value = target.node().value;

			if (value && value.substr(0, 1) == "=") {
				target.expression(value);
			} else {
				target.expression(null);
				target.value(value);
			}
		} else if (target.J4_cell) {
			// Missing caret hack
			target.value = " ";

			if (target.J4_cell.editing() && target.J4_cell.expression()) {
				target.value = target.J4_cell.expression();
			} else {
				target.value = target.J4_cell.value();
			}
		}
	}

	function onMouseDown(event) {
		var target = event.target || event.srcElement,
			cell = target.J4_cell;

		if (cell && ! cell.editing() && event.button == 0) {
			this.isMouseDown = true;
			this.cellFrom = cell;
			this.cellTo = null;

			// Firefox D&D workaround
			event.preventDefault();
			target.focus();
		}
	}

	function onMouseUp(event) {
		this.isMouseDown = false;
		event.preventDefault();
	}

	function onMouseLeave(event) {
		var target = event.target || event.srcElement;

		if (target.classList.contains(Config.css.container)) {
			this.isMouseDown = false;
		}
	}

	function onMouseMove(event) {
		var target, cell;

		if (this.isMouseDown && this.cellFrom) {
			target = event.target || event.srcElement;
			cell = target.J4_cell;

			if (cell && cell !== this.cellTo) {
				if (cell !== this.cellFrom) {
					this.model.select(this.cellFrom, cell);
				} else {
					this.model.select(cell);
				}

				event.preventDefault();
				this.cellTo = cell;
			}
		}
	}

	function onClick(event) {
		var target = event.target || event.srcElement,
			cell = target.J4_cell;

		if (cell && event.button == 0) {
			this.model.select(cell);

			if (event.type === "dblclick" && ! cell.editing()) {
				cell.editing(true);
				update(target);
			}
		}
	}

	function onKeyPress(event) {
		var target = event.target || event.srcElement,
			cell = target.J4_cell,
			key = event.which;

		if (cell && ! cell.editing() && key != 0) {
			target.value = String.fromCharCode(key);
			this.model.select(cell);
			cell.editing(true);

			event.preventDefault();
		}
	}

	function onKeyDown(event) {
		var target = event.target || event.srcElement,
			cell = target.J4_cell,
			key = event.keyCode,
			column, row, value;

		if (cell) {
			if (cell.editing()) {
				if (key == KEY_ESCAPE) {
					// Discard changes
					target.value = cell.value();
					cell.editing(false);
				} else if (key == KEY_ENTER) {
					// Save changes
					event.preventDefault();
					cell.editing(false);
					update(cell);
				}
			} else {
				if (isArrowKey(key)) {
					// Move (or expand/shrink) the selection
					cell = event.shiftKey && this.cellTo || this.cellFrom;

					if (cell) {
						// Calculate position of the new cell
						column = cell.column() + (key == KEY_LEFT && -1 || key == KEY_RIGHT && 1 || 0);
						row = cell.row() + (key == KEY_UP && -1 || key == KEY_DOWN && 1 || 0);
					}

					// Get the new cell or stop if a boundary was reached
					cell = this.model.cell(row, column) || cell;

					if (cell) {
						target = cell.node();

						if (event.shiftKey) {
							// Expand/shrink selected range
							this.model.select(this.cellFrom, cell);
							this.cellTo = cell;
						} else {
							// Move selection
							target.focus();
							this.model.select(cell);

							this.cellFrom = cell;
							this.cellTo = null;
						}
					}

					event.preventDefault();
				} else if (key == KEY_BACKSPACE || key == KEY_DELETE) {
					// Clear expression and value of each selected cell
					this.model.selectionExpression(null).selectionValue(null);
					event.preventDefault();
				} else if (key == KEY_ESCAPE) {
					this.model.selectionClear();
					event.preventDefault();
				}
			}
		}
	}

	function onLostFocus(event) {
		var target = event.target || event.srcElement,
			cell = target.J4_cell,
			editing;

		if (cell) {
			editing = cell.editing();
			cell.editing(false);
			target.blur();

			if (editing) {
				update(cell);
			}
		}

		this.model.selectionClear();
	}

	return {
		init: init
	};
})();

/*
 * END CONTROLLER
 *
*/

/*
 * API
 *
*/

// Create spreadsheet container that can be appended to the DOM tree.
// Container has attached and initialized it's controller.
function create(numRows, numColumns) {
	var spreadsheet = new Spreadsheet(numRows, numColumns),
		container = document.createElement("div"),
		table = document.createElement("table"),
		tbody = document.createElement("tbody"),
		input, text, tr, th, td,
		i, j;

	// Container
	container.classList.add(Config.css.container);
	container.appendChild(table);

	// Table
	table.classList.add(Config.css.spreadsheet);
	table.appendChild(tbody);

	// Header
	tr = table.createTHead().insertRow();
	th = document.createElement("th");

	th.classList.add(Config.css.header, Config.css.emptyHeader);
	tr.appendChild(th);

	for (i = 0; i < numColumns; i += 1) {
		text = document.createTextNode(encodeColumn(i));
		th = document.createElement("th");

		th.classList.add(Config.css.header, Config.css.columnHeader);
		th.appendChild(text);
		tr.appendChild(th);
	}

	// Body
	for (i = 0; i < numRows; i += 1) {
		text = document.createTextNode(i + 1);
		th = document.createElement("th");
		tr = tbody.insertRow();

		th.classList.add(Config.css.header, Config.css.rowHeader);
		th.appendChild(text);
		tr.appendChild(th);

		for (j = 0; j < numColumns; j += 1) {
			input = document.createElement("input");
			td = tr.insertCell();

			input.classList.add(Config.css.cellValue);
			input.setAttribute("readonly", "readonly");
			input.setAttribute("type", "text");

			td.classList.add(Config.css.cell);
			td.appendChild(input);
		}
	}

	Controller.init(container, spreadsheet);
	return container;
}

// Appends the spreadsheet container to parent of the last script
// in the DOM tree.
function spawn(numRows, numColumns) {
	var scripts = document.getElementsByTagName("script"),
		container = create(numRows, numColumns);

	scripts[scripts.length - 1].parentNode.appendChild(container);
}

/*
 * END API
 *
*/

// Return public API
return {
	create: create,
	spawn: spawn
};

})();
