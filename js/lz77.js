// Modified from: https://github.com/whoughton/lz77
(function(){
	'use strict';

	// - - -
	// # LZ77
	// ** A minimal LZ77 [de]compressor **
	var LZ77 = function(){
		// ##### Private Variables
		var
			self = {}
		,	settings = {
			refPrefix: '`',
			refIntBase: 96,
			refIntFloorCode: ' '.charCodeAt(0),
			refIntCeilCode: undefined,
			maxStringDistance: undefined,
			minStringLength: 5,
			maxStringLength: undefined,
			defaultWindow: 144,
			maxWindow: undefined,
			windowLength: undefined
		}
		;

		// ##### Public Variables

		// ##### Private Methods
		var setup = function(params) {
			params = params || {};
			settings = extend(settings, params);
	
			settings.refIntCeilCode = settings.refIntFloorCode + settings.refIntBase - 1;
			settings.maxStringDistance = Math.pow(settings.refIntBase, 2) - 1;
			settings.maxStringLength = Math.pow(settings.refIntBase, 1) - 1 + settings.minStringLength;
			settings.maxWindow = settings.maxStringDistance + settings.minStringLength;
		};

		// #### each()
		// >`@param obj [collection]` our source collection  
		// >`@param iterator [function]` the function that will be called for each element in the collection  
		// >`@param context [object]` the context our iterator should operate within  
		//
		// essentially copied from underscore.js
		var each = function(obj, iterator, context) {
			var breaker = {};
			if (obj === null) { return; }
			if (Array.prototype.forEach && obj.forEach === Array.prototype.forEach) {
				obj.forEach(iterator, context);
			} else if (obj.length === +obj.length) {
				for (var i = 0, l = obj.length; i < l; i++) {
					if (iterator.call(context, obj[i], i, obj) === breaker) { return; }
				}
			} else {
				for (var key in obj) {
					if (obj.hasOwnProperty(key)) {
						if (iterator.call(context, obj[key], key, obj) === breaker) { return; }
					}
				}
			}
		};

		// #### extend()
		// >`@param obj [object]` our destination object  
		// >`@params * [object(s)]` objects that will overwrite the data in the destination object, in order  
		//
		// essentially copied from underscore.js
		var extend = function(obj) {
			each(Array.prototype.slice.call(arguments, 1), function(source) {
				if (source) {
					for (var prop in source) {
						obj[prop] = source[prop];
					}
				}
			});
			return obj;
		};

		var encodeRefInt = function(value, width) {
			if ((value >= 0) && (value < (Math.pow(settings.refIntBase, width) - 1))) {
				var encoded = '';
				while (value > 0) {
					encoded = (String.fromCharCode((value % settings.refIntBase) + settings.refIntFloorCode)) + encoded;
					value = Math.floor(value / settings.refIntBase);
				}
				var missingLength = width - encoded.length;
				var i = 0;
				for (; i < missingLength; i++) {
					encoded = String.fromCharCode(settings.refIntFloorCode) + encoded;
				}
				return encoded;
			} else {
				throw 'Reference int out of range: ' + value + ' (width = ' + width + ')';
			}
		};

		var encodeRefLength = function(length) {
			return encodeRefInt(length - settings.minStringLength, 1);
		};

		var decodeRefInt = function(data, width) {
			var 
				value = 0
			,	charCode
			,	i = 0
			;
			for (; i < width; i++) {
				value *= settings.refIntBase;
				charCode = data.charCodeAt(i);
				if ((charCode >= settings.refIntFloorCode) && (charCode <= settings.refIntCeilCode)) {
					value += charCode - settings.refIntFloorCode;
				} else {
					throw 'Invalid char code in reference int: ' + charCode;
				}
			}
			return value;
		};

		var decodeRefLength = function(data) {
			return decodeRefInt(data, 1) + settings.minStringLength;
		};


		// ##### Public Methods

		// #### LZ77.compress()
		// >`@param source [string]` the source string we will be compressing  
		// >`@param params [object]` this allows us to overwrite defaults at call-time  
		//
		// This is our compression method, taking the input string (and allowing for call-time
		// paramters) and returning the compressed representation
		self.compress = function(source, params) {
			if (Object.prototype.toString.call(source) !== '[object String]') { throw 'Input must be a string'; }
			
			setup(params);

			var windowLength = settings.windowLength || settings.defaultWindow;
			if (windowLength > settings.maxWindow) { throw 'Window length too large'; }

			var
				compressed = ''
			,	pos = 0
			,	lastPos = source.length - settings.minStringLength
			;
	
			while (pos < lastPos) {
				var
					searchStart = Math.max(pos - windowLength, 0)
				,	matchLength = settings.minStringLength
				,	foundMatch = false
				,	bestMatch = {
					distance: settings.maxStringDistance,
					length: 0
				}
				,	newCompressed = null
				,	isValidMatch
				,	realMatchLength
				;
		
				while ((searchStart + matchLength) < pos) {
					isValidMatch = ((source.substr(searchStart, matchLength) === source.substr(pos, matchLength)) && (matchLength < settings.maxStringLength));
					if (isValidMatch) {
						matchLength++;
						foundMatch = true;
					} else {
						realMatchLength = matchLength - 1;
						if (foundMatch && (realMatchLength > bestMatch.length)) {
							bestMatch.distance = pos - searchStart - realMatchLength;
							bestMatch.length = realMatchLength;
						}
						matchLength = settings.minStringLength;
						searchStart++;
						foundMatch = false;
					}
				}
		
				if (bestMatch.length) {
					newCompressed = settings.refPrefix + encodeRefInt(bestMatch.distance, 2) + encodeRefLength(bestMatch.length);
					pos += bestMatch.length;
				} else {
					if (source.charAt(pos) !== settings.refPrefix) {
						newCompressed = source.charAt(pos);
					} else {
						newCompressed = settings.refPrefix + settings.refPrefix;
					}
					pos++;
				}
				compressed += newCompressed;
			}

			return compressed + source.slice(pos).replace(/`/g, '``');
		};

		// #### LZ77.decompress()
		// >`@param source [string]` the source string of compressed data  
		// >`@param params [object]` this allows us to overwrite defaults at call-time  
		//
		// decompression method, taking the compressed data (as a string, and allowing for
		// call-time paramters) and returning the decompressed data
		self.decompress = function(source, params) {
			if (Object.prototype.toString.call(source) !== '[object String]') { throw 'Input must be a string'; }
			
			var
				decompressed = ''
			,	pos = 0
			,	currentChar
			,	nextChar
			,	distance
			,	length
			;
	
			setup(params);
			while (pos < source.length) {
				currentChar = source.charAt(pos);
				if (currentChar !== settings.refPrefix) {
					decompressed += currentChar;
					pos++;
				} else {
					nextChar = source.charAt(pos + 1);
					if (nextChar !== settings.refPrefix) {
						distance = decodeRefInt(source.substr(pos + 1, 2), 2);
						length = decodeRefLength(source.charAt(pos + 3));
						decompressed += decompressed.substr(decompressed.length - distance - length, length);
						pos += settings.minStringLength - 1;
					} else {
						decompressed += settings.refPrefix;
						pos += 2;
					}
				}
			}
			return decompressed;
		};

		return self;
	}();

	if (typeof define !== 'undefined' && define.amd) {				// requirejs/amd
		define([], function() { return LZ77; });
	} else if (typeof module !== 'undefined' && module.exports) {	// node
		module.exports = LZ77;
	} else if (window !== undefined) {
		window.LZ77 = LZ77;
	}
})();