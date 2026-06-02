sap.ui.define(["sap/ui/core/mvc/Controller", "sap/ui/core/routing/History", "sap/m/MessageBox"], function (e, t, r) {
	"use strict";
	var a = "com.isa.presupuesto.zplanconspre";
	return e.extend("com.isa.presupuesto.zplanconspre.controller.BaseController", {
		getRouter: function () {
			return this.getOwnerComponent().getRouter()
		},
		getModel: function (e) {
			return this.getView().getModel(e)
		},
		setModel: function (e, t) {
			return this.getView().setModel(e, t)
		},
		getResourceBundle: function () {
			return this.getOwnerComponent().getModel("i18n").getResourceBundle()
		},
		getDataManagerDAO: function () {
			return this.getOwnerComponent().oDataManagerDAO
		},
		onNavBack: function () {
			var e = t.getInstance().getPreviousHash();
			var r;
			if (sap.ushell && sap.ushell.Container) {
				r = sap.ushell.Container.getService("CrossApplicationNavigation")
			}
			if (e !== undefined || r && !r.isInitialNavigation()) {
				history.go(-1)
			} else {
				if (this.getView().getViewName() === a + ".view.List") {
					this.getRouter().navTo("RouteFilters", {}, true)
				} else {
					if (r) {
						r.toExternal({
							target: {
								semanticObject: "Shell",
								action: "home"
							}
						})
					}
				}
			}
		},
		getModeApplication: function () {
			var e = "v";
			if (sap.ushell && sap.ushell.services) {
				var t = sap.ushell.services.AppConfiguration.getCurrentApplication().sShellHash;
				if (!t) {
					t = sap.ushell.services.AppConfiguration.getCurrentApplication().sFixedShellHash
				}
				var r = t.substring(t.indexOf("-") + 1);
				if (r.indexOf("?") >= 0) {
					r = r.substring(0, r.indexOf("?"))
				}
				switch (r) {
				case "display":
					e = "v";
					break;
				case "change":
					e = "e";
					break;
				case "create":
					e = "c";
					break;
				case "manage":
					e = "m";
					break;
				default:
					e = "n"
				}
			}
			return e
		},
		showGeneralError: function (e) {
			e = jQuery.extend({
				message: "",
				additionalData: "",
				oDataError: "",
				title: "",
				onClose: function () {}
			}, e);
			var t = "",
				a = "";
			if (e.message) {
				t = e.message
			} else {
				t = this.getResourceBundle().getText("technicalError")
			}
			if (e.additionalData) {
				a = e.additionalData
			}
			if (e.oDataError) {
				try {
					var n = jQuery.parseJSON(e.oDataError.responseText);
					if (n.error.code.indexOf("W5") >= 0 || n.error.code.indexOf("SY/530") || n.error.code.indexOf("KI") >= 0) {
						t = n.error.message.value;
						a = ""
					} else {
						a = n.error.message.value
					}
				} catch (t) {
					a = e.oDataError
				}
			}
			jQuery.sap.log.error(t);
			jQuery.sap.log.error(a);
			var i = e.title ? e.title : this.getResourceBundle().getText("errorTitleMessageBox");
			r.show(t, {
				icon: sap.m.MessageBox.Icon.ERROR,
				title: i,
				onClose: e.onClose,
				details: a,
				actions: sap.m.MessageBox.Action.CLOSE
			})
		},
		isIntentSupported: function (e, t, r) {
			var a = sap.ushell.Container.getService("CrossApplicationNavigation");
			var n = a && a.hrefForExternal({
				target: {
					semanticObject: "Reservation",
					action: e
				}
			}) || "";
			if (a) {
				a.isIntentSupported([n]).done(function (e) {
					if (e[n].supported === true) {
						r(true)
					} else {
						r(false)
					}
				}).fail(function () {
					r(false)
				})
			} else {
				r(false)
			}
		},
		navToAction: function (e, t) {
			var r = sap.ushell.Container.getService("CrossApplicationNavigation");
			var a = "";
			if (e === "display") {
				a = "&/view/" + t
			} else if (e === "change") {
				a = "&/edit/" + t
			} else if (e === "create") {
				a = "&/new"
			}
			var n = r && r.hrefForExternal({
				target: {
					shellHash: "Reservation-" + e + a
				}
			}) || "";
			var i = {
				target: {
					shellHash: n
				}
			};
			this.isIntentSupported(e, t, function (e) {
				if (e) {
					r.toExternal(i)
				}
			})
		},
		_closeWindow: function () {
			window.close()
		}
	})
});