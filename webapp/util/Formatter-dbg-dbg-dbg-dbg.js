sap.ui.define([], function () {
	"use strict";

	return {
		/**
		 * Se formatea string a fecha
		 * @param {String} sDate Fecha en string
		 * @param {Object} oContext Contexto de la vista. Opcional. Se env?a cuando se llama el metodo desde el controlador de una vista
		 * @return {Date} Fecha
		 * */
		formatDateFromString: function (sDate, oContext) {
			var oResoucerBundle;
			if (!oContext) {
				oResoucerBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
			} else {
				oResoucerBundle = oContext.getModel("i18n").getResourceBundle();
			}
			var sFormat = oResoucerBundle.getText("configFormatDatePicker");
			var oDateFormat = sap.ui.core.format.DateFormat.getTimeInstance({
				pattern: sFormat
			});
			var dFormatedDate = new Date(oDateFormat.parse(sDate).getTime());
			return dFormatedDate;
		},
		/**
		 * Se formatea fecha a string
		 * @param {Date} dDate Fecha a convertir
		 * @return {String} Fecha en el formato definido en el i18n configFormatDatePicker
		 * */
		formatDateToString: function (dDate, oContext) {
			var oResoucerBundle;
			if (!oContext) {
				oResoucerBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
			} else {
				oResoucerBundle = oContext.getModel("i18n").getResourceBundle();
			}
			if (dDate) {
				var iTime;
				if (typeof dDate === "string" && dDate.startsWith("/Date(")) {
					iTime = dDate.replace("/Date(", "").replace(")/", "");
					if (!isNaN(iTime)) {
						iTime = parseInt(iTime, 10);
					}
				} else {
					iTime = dDate.getTime();
				}
				if (!isNaN(iTime)) {
					var sFormat = oResoucerBundle.getText("configFormatDatePicker");
					var oDateFormat = sap.ui.core.format.DateFormat.getTimeInstance({
						pattern: sFormat
					});
					var sTimeOffset = new Date(0).getTimezoneOffset() * 60 * 1000;
					var sFormatedDate = oDateFormat.format(new Date(iTime + sTimeOffset));
					return sFormatedDate;
				}
			}
			return "";
		},

		formatDateTimeToString: function (dDate, dTime, oContext) {
			var oResoucerBundle;
			if (!oContext) {
				oResoucerBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
			} else {
				oResoucerBundle = oContext.getModel("i18n").getResourceBundle();
			}
			if (dDate) {
				var iTime;
				if (typeof dDate === "string" && dDate.startsWith("/Date(")) {
					iTime = dDate.replace("/Date(", "").replace(")/", "");
					if (!isNaN(iTime)) {
						iTime = parseInt(iTime, 10);
					}
				} else {
					iTime = dDate.getTime();
				}
				if (dTime) {
					iTime = iTime + dTime.ms;
				}
				if (!isNaN(iTime)) {
					var sFormat = oResoucerBundle.getText("configFormatDateTimePicker");
					var oDateFormat = sap.ui.core.format.DateFormat.getTimeInstance({
						pattern: sFormat
					});
					var sTimeOffset = new Date(0).getTimezoneOffset() * 60 * 1000;
					var sFormatedDate = oDateFormat.format(new Date(iTime + sTimeOffset));
					return sFormatedDate;
				}
			}
			return "";
		},

		formatTimeToString: function (dTime, oContext) {
			var oResoucerBundle;
			if (!oContext) {
				oResoucerBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
			} else {
				oResoucerBundle = oContext.getModel("i18n").getResourceBundle();
			}

			var iTime;
			if (dTime) {
				iTime = dTime.ms;
			}
			if (!isNaN(iTime)) {
				var sFormat = oResoucerBundle.getText("configFormatTimePicker");
				var oDateFormat = sap.ui.core.format.DateFormat.getTimeInstance({
					pattern: sFormat
				});
				var sTimeOffset = new Date(0).getTimezoneOffset() * 60 * 1000;
				var sFormatedDate = oDateFormat.format(new Date(iTime + sTimeOffset));
				return sFormatedDate;
			}
			return "";
		},

		formatDateToDateUTC: function (dDate, oContext) {
			var oResoucerBundle,
				dResult;
			if (!oContext) {
				oResoucerBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
			} else {
				oResoucerBundle = oContext.getModel("i18n").getResourceBundle();
			}
			if (dDate) {
				var iTime;
				if (typeof dDate === "string" && dDate.startsWith("/Date(")) {
					iTime = dDate.replace("/Date(", "").replace(")/", "");
					if (!isNaN(iTime)) {
						iTime = parseInt(iTime, 10);
					}
				} else {
					iTime = dDate.getTime();
				}
				if (!isNaN(iTime)) {
					var iTimeOffset = new Date(0).getTimezoneOffset() * 60 * 1000;
					dResult = new Date(iTime + iTimeOffset);
				}
			}
			return dResult;
		},

		formatDateTimeToDateUTC: function (dDate, dTime, oContext) {
			var oResoucerBundle,
				dResult;
			if (!oContext) {
				oResoucerBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
			} else {
				oResoucerBundle = oContext.getModel("i18n").getResourceBundle();
			}
			if (dDate) {
				var iTime;
				if (typeof dDate === "string" && dDate.startsWith("/Date(")) {
					iTime = dDate.replace("/Date(", "").replace(")/", "");
					if (!isNaN(iTime)) {
						iTime = parseInt(iTime, 10);
					}
				} else {
					iTime = dDate.getTime();
				}
				if (!isNaN(iTime)) {
					var iTimeOffset = new Date(0).getTimezoneOffset() * 60 * 1000;
					dResult = new Date(iTime + iTimeOffset);
				}

				if (dTime) {
					dResult = new Date(dResult.getTime() + dTime.ms);
				}
			}
			return dResult;
		},
		/**
		 * Se formatea fecha a string, cuando la fecha no se encuentra en UTC
		 * @param {Date} dDate Fecha a convertir
		 * @param {Object} oContext Contexto de la vista. Opcional. Se env?a cuando se llama el metodo desde el controlador de una vista
		 * @return {String} Fecha en el formato definido en el i18n configFormatDatePicker
		 * */
		formatLocalDate: function (dDate, oContext) {
			var oResoucerBundle;
			if (oContext) {
				oResoucerBundle = oContext.getModel("i18n").getResourceBundle();
			} else {
				oResoucerBundle = this.getOwnerComponent().getModel("i18n").getResourceBundle();
			}
			if (dDate) {
				var sFormat = oResoucerBundle.getText("configFormatDatePicker");
				var oDateFormat = sap.ui.core.format.DateFormat.getTimeInstance({
					pattern: sFormat
				});
				var sFormatedDate = oDateFormat.format(dDate);
				return sFormatedDate;
			}
			return "";
		},
		formatNumcToString: function (sValue) {
			var sReturn = "";
			if (sValue) {
				sReturn = parseInt(sValue, 10);
				if (sReturn === 0) {
					sReturn = "";
				}
			}

			return sReturn;
		},
		formatNumc: function (sValue) {
			if (sValue) {
				if (!isNaN(sValue)) {
					return parseInt(sValue, 10);
				}
			}
			return sValue;
		},
		formatNumber: function (sValue) {
			if (!sValue) {
				return "0";
			}
			var sNewValue = parseFloat(sValue).toFixed(2);
			if (sNewValue.length >= 2 && sNewValue.substring(sNewValue.length - 2) === "00") {
				sNewValue = parseInt(sNewValue, 10);
			}
			return sNewValue;
		},

		addLeftZeros: function (sValue, iLength) {
			var sNewValue = "";
			sValue = sValue.toString();
			if (sValue && sValue.length === iLength) {
				sNewValue = sValue;
			} else if (sValue && sValue.length < iLength) {
				sNewValue = sValue;
				while (sNewValue.length < iLength) {
					sNewValue = "0" + sNewValue;
				}
			}
			return sNewValue;
		},

		formatValueState: function (sValue) {
			if (!sValue) {
				return "None";
			} else {
				return sValue;
			}
		},

		formatPercent: function (sValue, bIsPercent) {
			var aFormatted = "";
			if (!sValue) {
				sValue = "0";
			}
			var fValue = parseFloat(sValue);
			if (bIsPercent) {
				fValue = fValue / 100;
				var aNumberFormatPercent = sap.ui.core.format.NumberFormat.getPercentInstance({
					maxFractionDigits: 2,
					minFractionDigits: 0
				});
				aFormatted = aNumberFormatPercent.format(fValue, "%");
			} else {
				var aNumberFormat = sap.ui.core.format.NumberFormat.getFloatInstance({
					maxFractionDigits: 2,
					minFractionDigits: 0
				});
				aFormatted = aNumberFormat.format(fValue);
			}
			return aFormatted;
		}
	};
});