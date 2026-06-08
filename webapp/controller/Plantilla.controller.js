//# sourceMappingURL=Plantilla.controller.js.map/* global xlsx: true */
/* global jszip: true*/
sap.ui.define([
	"./BaseController",
	"sap/ui/model/json/JSONModel",
	"sap/ui/core/Fragment",
	"sap/ui/model/Filter",
	"sap/ui/model/FilterOperator",
	"../util/Formatter",
	"sap/m/MessageBox"
], function (Controller, JSONModel, Fragment, Filter, FilterOperator, Formatter, MessageBox) {
	"use strict";
	var _namespace = "soc.ZPLANCONSPRE_SOCIEDADES";
	return Controller.extend(_namespace + ".controller.Plantilla", {
		Formatter: Formatter,
		onInit: function () {
			this._createViewModel();
			//	this.getRouter().getRoute("RoutePlantilla").attachPatternMatched(this._onViewMatched, this);
			//	this._onViewMatched();
			var sText = "Real\n" + "Periodo Anterior";
			this.getView().byId("iconRealText").setText(sText);
			this.Message = "";
			this.Destinos = [];
			

		},
		onSociedadCoVH: function () {
			var oView = this.getView();
			if (!this._pValueHelpDialogSociedadCo) {
				this._pValueHelpDialogSociedadCo = Fragment.load({
					id: oView.getId(),
					name: _namespace + ".fragment.SociedadCo",
					controller: this
				}).then(function (oValueHelpDialog) {
					oView.addDependent(oValueHelpDialog);
					return oValueHelpDialog;
				});
			}
			this._pValueHelpDialogSociedadCo.then(function (oValueHelpDialog) {
				oValueHelpDialog.open();
			}.bind(this));
		},
		onSociedadFiVH: function (oEvent) {
			var that = this;
			var oView = this.getView();
			if (!this._pValueHelpDialogSociedadFi) {
				this._pValueHelpDialogSociedadFi = Fragment.load({
					id: oView.getId(),
					name: _namespace + ".fragment.SociedadFi",
					controller: this
				}).then(function (oValueHelpDialog) {
				
					oView.addDependent(oValueHelpDialog);
					return oValueHelpDialog;
				});
			}
			this._pValueHelpDialogSociedadFi.then(function (oValueHelpDialog) {
				oValueHelpDialog.open();
			}.bind(this));

		},
		onCecoVH: function () {
			var oView = this.getView();
			var oFilter = [new Filter("Kokrs", FilterOperator.EQ, oView.byId("inpSociedadCo").getValue())];
			oFilter.push(new Filter("Bukrs", FilterOperator.EQ, oView.byId("inpSociedadFi").getValue()));

			oView.getModel().read("/CecoSet", {
				filters: oFilter,
				success: function (response) {
					var ceco = new JSONModel(response.results);
					oView.setModel(ceco, "CecoLocalModel");
				},
				error: function (error) {
					debugger;
				}
			});
			if (!this._pValueHelpDialogCeco) {
				this._pValueHelpDialogCeco = Fragment.load({
					id: oView.getId(),
					name: _namespace + ".fragment.CecoValueHelp",
					controller: this
				}).then(function (oValueHelpDialog) {
					oView.addDependent(oValueHelpDialog);
					return oValueHelpDialog;
				});
			}
			this._pValueHelpDialogCeco.then(function (oValueHelpDialog) {
				oValueHelpDialog.open();
			}.bind(this));
		},

		onSearchSociedadCo: function (oEvent) {
			var sValue = oEvent.getParameter("value").toUpperCase();
			var oFilter = new Filter("Kokrs", FilterOperator.EQ, sValue);
			var oBinding = oEvent.getParameter("itemsBinding");
			oBinding.filter([oFilter]);
		},
		onValueHelpSociedadCo: function (oEvent) {
			this._pValueHelpDialogSociedadFi = undefined;
			var oSelectedItem = oEvent.getParameter("selectedItem"),
				oInput = this.byId("inpSociedadCo");
			if (!oSelectedItem) {
				oInput.resetProperty("value");
				return;
			}
			
			var sociedadCO = new JSONModel({
				bukrs: this.getView().byId("inpSociedadCo").getValue().toUpperCase()
			})
		
			
			this.getView().setModel(sociedadCO,"SociedadCo");
			
			oInput.setValue(oSelectedItem.getTitle());

		},

		onUpload: function (oEvent) {
			that.oGlobalBusyDialog.setText("Importando documento excel. Aguarde por favor");
			that.oGlobalBusyDialog.open();
			this._import(oEvent.getParameter("files") && oEvent.getParameter("files")[0]);
		},

		_import: function (file) {
			var excelData = {};
			var excelDato = [];
			var oExcelModel = this.getView().getModel("excelModel");
			var oDatos = oExcelModel.getData();
			var excelDataJson = {};
			var newDateStart, newDateEnd;
			if (file && window.FileReader) {
				var reader = new FileReader();
				reader.onload = function (oEvent) {
					var data = oEvent.target.result;
					var workbook = XLSX.read(data, {
						type: 'binary'
					});
					workbook.SheetNames.forEach(function (sheetName) {
						// Here is your object for every sheet in workbook
						excelData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

						for (var i = 0; i < excelData.length; i++) {
							excelDataJson = {
								AssignmentName: excelData[i]["Asignacion"],
								RKOSTL: excelData[i]["Ceco"],
								VORNR: excelData[i]["Operacion"],
								UVORN: excelData[i]["Suboperacion"],
								RPRZNR: excelData[i]["Proceso"],
								RNPLNR: excelData[i]["Grafo"],
								RKSTR: excelData[i]["Objeto"],
								RAUFNR: excelData[i]["Orden"],
								LSTAR: excelData[i]["Cl Act"],
								ValidityStartDate: excelData[i]["Fecha Inicial"],
								ValidityEndDate: excelData[i]["Fecha Final"]
							};

							excelDato.push(excelDataJson);

						}
					});
					if (Object.keys(oDatos).length !== 0) {
						oDatos.forEach((Obj) => {
							excelDato.push(Obj);
						})
					}

					if (excelDato.length !== 0) {
						this.onEnabledButtonCheck();
						// this.onEnableButtonSubmit();
					}
					that.oGlobalBusyDialog.close();
					// Setting the data to the local model 
					oExcelModel.setData(excelDato);
					/*that.localModel.setData({
						items: excelData
					});
					that.localModel.refresh(true);*/

					oExcelModel.refresh(true);

				}.bind(this);
				reader.onerror = function (ex) {
					that.oGlobalBusyDialog.close();
					console.log(ex);
				}.bind(this);
				reader.readAsBinaryString(file);
			}
		},
		cambioFi: function () {
			// debugger;
		},
		onSearchSociedadFi: function (oEvent) {
			var sValue = oEvent.getParameter("value").toUpperCase() || this.getView().byId("inpSociedadCo").getValue().toUpperCase();
			var oFilter = new Filter("Bukrs", FilterOperator.EQ, sValue);
			var oBinding = oEvent.getParameter("itemsBinding");
			oBinding.filter([oFilter]);

		},

		onValueSociedadFiClose: function (oEvent) {
			var oSelectedItem = oEvent.getParameter("selectedItem"),
				oInput = this.byId("inpSociedadFi");

			if (!oSelectedItem) {
				oInput.resetProperty("value");
				return;
			}

			oInput.setValue(oSelectedItem.getTitle());
			this._pValueHelpDialogSociedadFi.then(function (oValueHelpDialog) {
				oValueHelpDialog.close();
			}.bind(this));

		},
		
		onSearchCeco: function (oEvent) {
			var sValue = oEvent.getParameter("value").toUpperCase();
			var oFilter = new Filter("Kostl", FilterOperator.Contains, sValue);
			var oBinding = oEvent.getParameter("itemsBinding");
			oBinding.filter([oFilter]);
		},

		onValueHelpCecoClose: function (oEvent) {
			var oSelectedItem = oEvent.getParameter("selectedItem"),
				oInput = this.byId("inpCeco");

			if (!oSelectedItem) {
				oInput.resetProperty("value");
				return;
			}
			oInput.setValue(oSelectedItem.getTitle());
			this._pValueHelpDialogSociedadCo.then(function (oValueHelpDialog) {
				oValueHelpDialog.close();
			}.bind(this));
		},
		onGo: async function (oEvent) {
			await this._onViewMatched();
			var oViewModel = this.getModel("viewModel");
			var aConsolidado = this.getModel("TotalesModel").getProperty("/List");
			var sCeco = oViewModel.getProperty("/Ceco").toUpperCase();
			var sComp = oViewModel.getProperty("/Compa").toUpperCase();
			var dIper = oViewModel.getProperty("/Iper");
			var dFper = oViewModel.getProperty("/Fper");
			var oInpCeco = this.byId("inpCeco");
			var oDpIper = this.byId("dpIper");
			var oDpFper = this.byId("dpFper");
			var sociedadCo = this.byId("inpSociedadCo").getValue().toUpperCase();;
			var sociedadFi = this.byId("inpSociedadFi").getValue().toUpperCase();;
			oInpCeco.setValueState("None");
			oDpIper.setValueState("None");
			oDpFper.setValueState("None");
			if (sCeco && dIper && dFper) {
				var sIper = Formatter.addLeftZeros((dIper.getMonth() + 1).toString(), 3);
				var sFper = Formatter.addLeftZeros((dFper.getMonth() + 1).toString(), 3);
				var sYear = dIper.getFullYear().toString();
				var sYear2 = dFper.getFullYear().toString();
				if (sYear !== sYear2) {
					this.showGeneralError({
						message: "No coinciden los años en los periodos seleccionados."
					});
					return;
				}
				if (dIper > dFper) {
					this.showGeneralError({
						message: "La fecha inicial debe ser menor a la final."
					});
					return;
				}
				var aItems = [];
				oViewModel.setProperty("/busy", true);
				this.sCeco = sCeco;
				this.sYear = sYear;
				this.sIper = sIper;
				this.sFper = sFper;
				var r = await this._getReal(sCeco, sYear, sIper, sFper);
				
				/*if ( r.results.length == 0){
					MessageBox.error("No se encontro  Ceco");
					oViewModel.setProperty("/busy", false);
					return;
				}
                    */
				oViewModel.setProperty("/real", r.results);
				var oView = this.getView(),
					oTable = oView.byId("tableReal");
				//oTable.setFixedRowCount(r.length);
				var oPersonigrama = await this._getPersonigrama(sCeco, sYear, sociedadCo, sociedadFi);
				var aDestinos = await this._getDestinos(sCeco, sYear, sIper, sFper, sociedadCo, sociedadFi);
				oViewModel.setProperty("/destino", aDestinos.results);
				this.Destinos = aDestinos.results;
				var aPlantilla = await this._getPlantilla(sCeco, sYear, sIper, sFper, sociedadCo, sociedadFi);
				var aActividades = await this._getActividades(this.sCeco, this.sYear, this.sIper, this.sFper, sociedadCo);

				

				oViewModel.setProperty("/actividades", aActividades.results);
				var that = this;
				// this._destinos.then(function (aDestinos) {
				// 	oViewModel.setProperty("/destino", aDestinos.results);
				// 	that.Destinos = aDestinos.results;
				// });
				// aActividades = aActividades.results;
				// for (var i = 0; i < aActividades.length; i++) {
				// 	var oItem = this._getnewLine();
				// 	oItem.Ceco = sCeco;
				// 	oItem.CecoName = this.getModel().getProperty("/CecoSet('" + sCeco + "')");
				// 	oItem.Actividad = aActividades[i].Lstar;
				// 	oItem.ActividadName = aActividades[i].Ltext;
				// 	for (var k = 1; k <= 12; k++) {
				// 		oItem['mes' + k] = 0;
				// 	}
				// 	aItems.push(oItem);
				// }
				var iIper = parseInt(sIper, 10);
				var iFper = parseInt(sFper, 10);
				for (var j = 1; j <= 12; j++) {
					if (j < iIper || j > iFper) {
						oViewModel.setProperty("/mes" + j + "Visible", false);
					} else {
						oViewModel.setProperty("/mes" + j + "Visible", true);
					}
				}
				aPlantilla = aPlantilla.results;
				if ( aPlantilla.length == 0){
					MessageBox.error("No se encontro registro en la tabla Responsables por Ceco");
					oViewModel.setProperty("/busy", false);
					return;
				}
				
				 for (var i = 0; i < aPlantilla.length; i++) {
					var oItem = this._getnewLine();
					oItem.Ceco = sCeco;
					oItem.CecoName = aPlantilla[i].LtextCeco;
				
				
				
					oItem.Actividad = aPlantilla[i].Lstar;
					
				
					oItem.ActividadName = aPlantilla[i].LtextAct;
					oItem.Destino = aPlantilla[i].Desti;
					oItem.TipoRecep = aPlantilla[i].TipoRecep.toUpperCase();
					
					switch (oItem.TipoRecep) {
					case "ORDEN":
						oItem.TipoRecepKey = "01";
						break;
					case "PROCESO EMPRESARIAL":
						oItem.TipoRecepKey = "02";
						break;
					case "CENTRO DE COSTO":
						oItem.TipoRecepKey = "03";
						break;
					default:
						oItem.TipoRecepKey = "00";
						break;
					}
					oItem.DestinoVS = oItem.Destino.length === 0 ? "Error" : "None";
					oItem.DestinoVs = aPlantilla[i].DestiDesc;
					for (var k = 1; k <= 12; k++) {
						var fValue = parseFloat(aPlantilla[i]['Mes' + k], 10);
						if (isNaN(fValue)) {
							iValue = 0;
						}
						oItem['mes' + k] = fValue;
					}
					aItems.push(oItem);
				}
				oViewModel.setProperty("/items", aItems);
				var oBoolean = false;
				if (aPlantilla.length > 0 && aPlantilla[0].Lyear !== sYear) {
					this.disableFields(oBoolean);
					MessageBox.error("Actualmente no se pueden registrar datos para este periodo ya que se encuentra cerrado");
				} else {
					oBoolean = true;
					this.disableFields(oBoolean);
				}
				if (!isNaN(oPersonigrama.NumEmpl)) {
					oViewModel.setProperty("/personigrama", oPersonigrama);
					var iNumEmpl = parseInt(oPersonigrama.NumEmpl, 10);
				}
				
		
				let totalMeses = {};


				for (let j = 1; j <= 12; j++) {
					totalMeses["mes" + j] = 0;
				}


				for (let i = 0; i < aItems.length; i++) {
					for (let j = 1; j <= 12; j++) {
						let key = "mes" + j;
						let valor = parseFloat(aItems[i][key]) || 0;

						if (valor !== 0) {
							totalMeses[key] += valor;
						}
					}
				}

	
				// for (let j = 1; j <= 12; j++) {
				// 	aConsolidado[1]["mes" + j] = oPersonigrama["Mes" + j];
				// 	//aConsolidado[2]["mes" + j] = totalMeses["mes" + j];

				// 	//total Capacidad(b)
				// 	aConsolidado[3]["mes" + j] = oPersonigrama["CapMaxMes" + j];
				// }


				aConsolidado[1]["Title"] = "Cantidad de empleados plan " + sYear;
				this.getModel("TotalesModel").setProperty("/List", aConsolidado);
				this.getModel("TotalesModel").refresh();
				
				this.onChangeInput();


				
				oViewModel.setProperty("/busy", false);

				oViewModel.refresh();
				/*.then(function (oData) {
					console.log(oData);
				}).catch(function (oError) {
					console.log(oError);
				});*/
				/*var aActivities = this.getModel("ActividadModel").getProperty("/List");
				for (var i = 0; i < 5; i++) {
					var oItem = this._getnewLine();
					oItem.Ceco = sCeco;
					oItem.CecoName = "sNameCeco";
					oItem.Actividad = aActivities[i].key;
					aItems.push(oItem);
				}
				oViewModel.setProperty("/items", aItems);
				oViewModel.refresh();*/
	
		
		//	buttonSave.setEnabled(true);
			 await this.validateCapMax(aPlantilla,oPersonigrama);


			} else {
				if (!sCeco) {
					oInpCeco.setValueState("Error");
				}
				if (!dIper) {
					oDpIper.setValueState("Error");
				}
				if (!dFper) {
					oDpFper.setValueState("Error");
				}
			}

			this.__calcularTotalesItems();
			this.__createPlanillaDatos(aConsolidado,oPersonigrama,true);
    	
		},

	__createPlanillaDatos: function(aConsolidado, oPersonigrama, bSilenciarMensajes = false) {
		let buttonSave = this.byId("buttonSave");
		var aItems = this.getModel("viewModel").getProperty("/items");
		let totalMeses = {};
		for (let j = 1; j <= 12; j++) {
			totalMeses["mes" + j] = 0;
		}

		for (let i = 0; i <= 4; i++) {
			if (!aConsolidado[i]) {
				aConsolidado[i] = {};
			}
		}

		for (let i = 0; i < aItems.length; i++) {
			if (aItems[i].Ceco === "Total") continue;

			for (let j = 1; j <= 12; j++) {
				let key = "mes" + j;
				let valor = parseFloat(aItems[i][key]) || 0;
				totalMeses[key] += valor;
			}
		}

		let bError = false;
		let totalGeneraTotalPlanActividades = 0;
		let totalCantidadEmpleadosPlan2026 = 0;
		let totaTotalHorasHombreMes = 0;
		let totalCapacidadB = 0;
		let totalPorcentajeDeCapacidad = 0;
		let btnactive = true;

		for (let j = 1; j <= 12; j++) {
			let key = "mes" + j;
			let mesValue = totalMeses[key] || 0;
			let empleados = parseFloat(oPersonigrama["Mes" + j]) || 0;
			let capacidad = parseFloat(oPersonigrama["CapMaxMes" + j]) || 0;

			aConsolidado[0][key] = mesValue;
			totalGeneraTotalPlanActividades += mesValue;

			aConsolidado[1][key] = empleados;
			totalCantidadEmpleadosPlan2026 += empleados;

			aConsolidado[2][key] = empleados !== 0 ? mesValue / empleados : 0;
			totaTotalHorasHombreMes += aConsolidado[2][key];

			aConsolidado[3][key] = capacidad;
			totalCapacidadB += capacidad;

			// aConsolidado[4][key] = capacidad !== 0 ? mesValue / capacidad : 0;
			aConsolidado[4][key] = capacidad !== 0 ? (mesValue / capacidad) * 100 : 0;

			totalPorcentajeDeCapacidad += aConsolidado[4][key];

			let sValueState = "Success";
			if (Math.round(mesValue) !== Math.round(capacidad)) {
				sValueState = "Error";
				bError = true;
			}

			if (Math.round(mesValue) > Math.round(capacidad)) {
				btnactive = false;
			}
			
			aConsolidado[0][key + "VS"] = sValueState;
			aConsolidado[4][key + "VS"] = sValueState;
			
		}

		aConsolidado[0]["total"] = totalGeneraTotalPlanActividades;
		aConsolidado[0]["promedio"] = totalGeneraTotalPlanActividades / 12;

		aConsolidado[1]["total"] = totalCantidadEmpleadosPlan2026;
		aConsolidado[1]["promedio"] = totalCantidadEmpleadosPlan2026 / 12;

		aConsolidado[2]["total"] = totaTotalHorasHombreMes;
		aConsolidado[2]["promedio"] = totaTotalHorasHombreMes / 12;

		aConsolidado[3]["total"] = totalCapacidadB;
		aConsolidado[3]["promedio"] = totalCapacidadB / 12;

		aConsolidado[4]["total"] = totalPorcentajeDeCapacidad / 12;
		aConsolidado[4]["promedio"] = totalPorcentajeDeCapacidad / 12;


		if (!bError) {
			this.getModel("viewModel").setProperty("/iconColorTab", "Positive");
			buttonSave.setEnabled(true);
		} else {
			this.getModel("viewModel").setProperty("/iconColorTab", "Negative");
			buttonSave.setEnabled(false);
		}
		if (!btnactive) {
		buttonSave.setEnabled(false);
		if (!bSilenciarMensajes) {
			MessageBox.error("Hay meses donde el porcentaje de capacidad es superior al 100%. No se permite continuar.");
		}
	}
	/*
	else if (bError) {
		buttonSave.setEnabled(true);
		if (!bSilenciarMensajes) {
			MessageBox.warning("Hay meses donde el porcentaje de capacidad es inferior al 100%.");
		}
	} 
	*/
	else {
		buttonSave.setEnabled(true);
	}



		this.getModel("TotalesModel").setProperty("/List", aConsolidado);
		this.getModel("TotalesModel").refresh();
	},

	__calcularTotalesItems: function () {
		var aItems = this.getModel("viewModel").getProperty("/items");
		var aTotales = [];
		var iTotalGeneral = 0;

		for (var i = 0; i < aItems.length; i++) {
			var iTotalLinea = 0;

			if (aItems[i].Ceco !== "Total") {
				for (var j = 1; j <= 12; j++) {
					let mesKey = "mes" + j;

					if (
						aItems[i][mesKey] === undefined ||
						aItems[i][mesKey] === "" ||
						isNaN(parseFloat(aItems[i][mesKey]))
					) {
						aItems[i][mesKey] = 0;
					} else {
						aItems[i][mesKey] = parseFloat(aItems[i][mesKey]);
					}

					if (!aTotales[mesKey]) {
						aTotales[mesKey] = 0;
					}

					aTotales[mesKey] += aItems[i][mesKey];
					iTotalLinea += aItems[i][mesKey];
				}

				let fPromedio = iTotalLinea / 12;
				aItems[i]["total"] = iTotalLinea;
				aItems[i]["promedio"] = fPromedio;

				iTotalGeneral += iTotalLinea;
			}
		}

		var oTotalRow = aItems.find(function(item) {
			return item.Ceco === "Total";
		});

		if (!oTotalRow) {
			oTotalRow = this._getnewLine();
			oTotalRow.Ceco = "Total";
			aItems.push(oTotalRow);
		}

		for (var k = 1; k <= 12; k++) {
			oTotalRow["mes" + k] = aTotales["mes" + k];
		}

		oTotalRow["total"] = iTotalGeneral;
		oTotalRow["promedio"] = 0;

		for (var n = 0; n < aItems.length; n++) {
			if (aItems[n].Ceco !== "Total") {
				oTotalRow["promedio"] += aItems[n]["promedio"];
			}
		}

		this.getModel("viewModel").setProperty("/items", aItems);
		this.getModel("viewModel").refresh();
	},


		uploadFile: function (e) {
			this._import(e.getParameter("files") && e.getParameter("files")[0]);
		},
		_import: function (file) {
			var that = this;
			var excelData = {};
			if (file && window.FileReader) {
				var reader = new FileReader();
				reader.onload = function (e) {
					var data = e.target.result;
					var workbook = XLSX.read(data, {
						type: 'binary'
					});
					workbook.SheetNames.forEach(function (sheetName) {
						// Here is your object for every sheet in workbook
						excelData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

					});
					// Setting the data to the local model 
					that.localModel.setData({
						items: excelData
					});
					that.localModel.refresh(true);
				};
				reader.onerror = function (ex) {
					console.log(ex);
				};
				reader.readAsBinaryString(file);
			}
		},
		onAddItem: function (oEvent) {
			var oView = this.getView();
			var oViewModel = this.getModel("viewModel");
			if (!this._pDialogAddItem) {
				this._pDialogAddItem = Fragment.load({
					id: oView.getId(),
					name: _namespace + ".fragment.NewItemDialog",
					controller: this
				}).then(function (oValueHelpDialog) {
					oView.addDependent(oValueHelpDialog);
					return oValueHelpDialog;
				});
				// this._actividades = this._getActividades(this.sCeco, this.sYear, this.sIper, this.sFper);
				// this._actividades.then(function (aActividades) {
				// 	oViewModel.setProperty("/actividades", aActividades.results);
				// });
			}
			this._pDialogAddItem.then(function (oValueHelpDialog) {
				oValueHelpDialog.open();
				oValueHelpDialog.setBusy(false);
			}.bind(this));
			oViewModel.setProperty("/newItem/actividad", "");
			/*var oViewModel = this.getModel("viewModel");
			var aItems = oViewModel.getProperty("/items");
			aItems.push(this._getnewLine());
			oViewModel.setProperty("/items", aItems);
			oViewModel.refresh();*/
		},

	onDeleteItem: function (oEvent) {
		var oViewModel = this.getModel("viewModel");
		var sPath = oEvent.getParameter("item").getBindingContext("viewModel").getPath();
		var aItems = oViewModel.getProperty("/items");
		var iIndex = parseInt(sPath.replaceAll("/items/", ""));

		// Evitar eliminar si el índice no es válido o es la fila "Total"
		if (isNaN(iIndex) || !aItems[iIndex] || aItems[iIndex].Ceco === "Total") {
			return;
		}

		aItems.splice(iIndex, 1);
		oViewModel.setProperty("/items", aItems);
		oViewModel.refresh();

		this.__calcularTotalesItems();

		var aConsolidado = this.getModel("TotalesModel").getProperty("/List");
		var oPersonigrama = this.getModel("viewModel").getProperty("/personigrama");

		this.__createPlanillaDatos(aConsolidado, oPersonigrama, true);
	},

		onCloseNewItem: function (oEvent) {
			this._pDialogAddItem.then(function (oValueHelpDialog) {
				oValueHelpDialog.close();
			}.bind(this));
		},
		onConfirmNewItem: function (oEvent) {
			var oViewModel = this.getModel("viewModel");
			var sActividad = oViewModel.getProperty("/newItem/actividad");
			if (sActividad) {
				var aItems = oViewModel.getProperty("/items");
				//if (aItems.length > 0) {
				aItems.splice(aItems.length - 1);
				var oNewItem = this._getnewLine();
				oNewItem.Ceco = aItems[0].Ceco;
				oNewItem.CecoName = aItems[0].CecoName;
				oNewItem.Actividad = sActividad;
				var oLength = sActividad.length + 3;
				oNewItem.ActividadName = this.byId("cmbActividadesNew").getSelectedItem().getText().substring(oLength);
				oNewItem.DestinoVS = "Error";
				aItems.push(oNewItem);
				oViewModel.setProperty("/items", aItems);
				oViewModel.refresh();
				//}
				this.onChangeInput();
			}
			this._pDialogAddItem.then(function (oValueHelpDialog) {
				oValueHelpDialog.close();
			}.bind(this));
		},



		onChangeInput: function (oEvent) {
			if (!oEvent || !oEvent.getSource) return;

			var sValue = oEvent.getParameter("newValue");
			var oInput = oEvent.getSource();
			var oContext = oInput.getBindingContext("viewModel");
			var sPath = oInput.getBindingPath();

			if (!sValue || sValue.trim() === "") {
				sValue = "0";
				oInput.setValue(sValue);
			}

			if (oContext && sPath) {
				oContext.getModel().setProperty(oContext.getPath() + "/" + sPath, sValue);
			}

			var oDeletedModel = this.getView().getModel("valuesDelet");
			if (oDeletedModel) {
				var aEliminados = oDeletedModel.getProperty("/valuesDelet") || [];
				var oItem = oContext.getObject();

				var iIndex = aEliminados.findIndex(function (e) {
					return (
						e.Ceco === oItem.Kostl &&
						e.Actividad === oItem.Lstar &&
						e.Destino === oItem.Desti
					);
				});

				if (iIndex !== -1) {
					aEliminados[iIndex]["Mes" + oItem.Mes] = sValue;
					oDeletedModel.setProperty("/valuesDelet", aEliminados);
				}
			}

			var aConsolidado = this.getModel("TotalesModel").getProperty("/List");
			var oPersonigrama = this.getModel("viewModel").getProperty("/personigrama");

			this.__calcularTotalesItems();
			this.__createPlanillaDatos(aConsolidado, oPersonigrama, false);
		},



onSaveData: function () {
    var oViewModel = this.getModel("viewModel");
    var aConsolidado = this.getModel("TotalesModel").getProperty("/List");
    var bError = this._checkDatosPlantilla();
    var oPersonigrama = oViewModel.getProperty("/personigrama");
    var sMessage = "Se enviaron los datos al servidor. Desea continuar?";
    var sVersion = oPersonigrama.Version;

    if (bError) {
        sMessage = "Hay meses donde el porcentaje de capacidad no es el 100%. Sin embargo el sistema permitirá guardar. ¿Desea continuar?";
    }

    if (sVersion === "V2" && bError) {
        MessageBox.error("Hay meses con errores. Con la versión 2 del personigrama no se puede almacenar con diferencias en los cálculos.");
    } else {
        MessageBox.confirm(sMessage, {
            onClose: function (oAction) {
                if (oAction === "OK") {
                    var oData = this._getData();
                    if (this.validaDatos(oData)) {
                       // oData.Items.splice(oData.Items.length - 1);

                        // var oDeletedModel = this.getView().getModel("valuesDelet");
                        // if (oDeletedModel) {
                        //     var aEliminados = oDeletedModel.getProperty("/valuesDelet") || [];

                        //     var aItemsConvertidos = aEliminados
                        //         .filter(function (item) {
                        //             return !oData.Items.some(function (dataItem) {
                        //                 return (
                        //                     dataItem.Kostl === item.Ceco &&
                        //                     dataItem.Lstar === item.Actividad &&
                        //                     dataItem.Desti === item.Destino
                        //                 );
                        //             });
                        //         })
                        //         .map(function (item) {
                        //             return {
                        //                 Kostl: item.Ceco,
                        //                 Gjahr: this.sYear,
                        //                 Lstar: item.Actividad,
                        //                 Desti: item.Destino,
                        //                 Comp: oData.Items[0].Comp,
                        //                 TipoRecep: item.TipoRecep,
                        //                 Mes1: item.Mes1 || "0",
                        //                 Mes2: item.Mes2 || "0",
                        //                 Mes3: item.Mes3 || "0",
                        //                 Mes4: item.Mes4 || "0",
                        //                 Mes5: item.Mes5 || "0",
                        //                 Mes6: item.Mes6 || "0",
                        //                 Mes7: item.Mes7 || "0",
                        //                 Mes8: item.Mes8 || "0",
                        //                 Mes9: item.Mes9 || "0",
                        //                 Mes10: item.Mes10 || "0",
                        //                 Mes11: item.Mes11 || "0",
                        //                 Mes12: item.Mes12 || "0"
                        //             };
                        //         }.bind(this));

                        //    oData.Items = oData.Items.concat(aItemsConvertidos);
                        // }

                        oViewModel.setProperty("/busy", true);

                        this.getModel().create("/PlantillaHeaderSet", oData, {
                            success: function (oData) {
                                MessageBox.success("Se almacenaron exitosamente los datos.");
                                oViewModel.setProperty("/busy", false);
                            }.bind(this),
                            error: function (oError) {
                                var mensaje = JSON.parse(oError.responseText).error.message.value;
                                this.showGeneralError({
                                    oDataError: mensaje || oError
                                });
                                oViewModel.setProperty("/busy", false);
                            }.bind(this)
                        });
                    } else {
                        MessageBox.error(this.Message);
                    }
                }
            }.bind(this)
        });
    }
},





		validaDatos: function (oData, Message) {
			var valida = oData.Items.length - 1;
			for (var a = 0; a < oData.Items.length; a++) {
				if (valida !== a) {
					for (var b = 0; b < oData.Items.length; b++) {
						if ((oData.Items[a].Desti === oData.Items[b].Desti) &&
							(oData.Items[a].DestiDesc === oData.Items[b].DestiDesc) &&
							(oData.Items[a].Gjahr === oData.Items[b].Gjahr) &&
							(oData.Items[a].Kostl === oData.Items[b].Kostl) &&
							(oData.Items[a].Lstar === oData.Items[b].Lstar) &&
							(oData.Items[a].TipoRecep === oData.Items[b].TipoRecep) &&
							(a !== b)) {
							this.Message = "Existen registros duplicados, por favor corregir la informacion";
							return false;
						}
					}
					if (oData.Items[a].Desti === '' || oData.Items[a].Desti === null || oData.Items[a].Desti === "0") {
						this.Message = "No se puede guardar registros sin asignar un destino";
						 return false;
					}
				
				
					if (oData.Items[a].Mes1.substring(0, 1) === '-' || oData.Items[a].Mes2.substring(0, 1) === '-' || oData.Items[a].Mes3.substring(0,
							1) === '-' ||
						oData.Items[a].Mes4.substring(0, 1) === '-' || oData.Items[a].Mes5.substring(0, 1) === '-' || oData.Items[a].Mes6.substring(0, 1) ===
						'-' ||
						oData.Items[a].Mes7.substring(0, 1) === '-' || oData.Items[a].Mes8.substring(0, 1) === '-' || oData.Items[a].Mes9.substring(0, 1) ===
						'-' ||
						oData.Items[a].Mes10.substring(0, 1) === '-' || oData.Items[a].Mes11.substring(0, 1) === '-' || oData.Items[a].Mes12.substring(0,
							1) === '-') {
						this.Message = "No se aceptan valores negativos";
						return false;
					}
				}
			}
			return true;
		},
		onDestinoVH: function (oEvent) {
			var oView = this.getView();
			var oModel = this.getModel("viewModel");
			let oDestino = [];
			var aItem = oEvent.getSource().getBindingContext("viewModel").getObject();
			if (aItem.TipoRecep.toUpperCase() === "PROCESO EMPRESARIAL") {
				oDestino = this.Destinos.filter(function (element) {
					return element.Desti === aItem.Actividad;
				})
			}
			else {
				oDestino = this.Destinos.filter(function (element) {
					return element.TipoRecep.toUpperCase()  === aItem.TipoRecep.toUpperCase() ;
				})
			}
			oModel.setProperty("/destino", oDestino);

			this._inputDestino = oEvent.getSource();
			if (!this._pValueHelpDialogDestino) {
				this._pValueHelpDialogDestino = Fragment.load({
					id: oView.getId(),
					name: _namespace + ".fragment.DestinoValueHelp",
					controller: this
				}).then(function (oValueHelpDialog) {
					oView.addDependent(oValueHelpDialog);
					oValueHelpDialog.setBusy(true);
					// this._destinos.then(function (aDestinos) {
					oValueHelpDialog.setBusy(false);
					// });
					return oValueHelpDialog;
				}.bind(this));
			}
			this._pValueHelpDialogDestino.then(function (oValueHelpDialog) {
				oValueHelpDialog.open();
			}.bind(this));
		},
		onSearchDestino: function (oEvent) {
			var sValue = oEvent.getParameter("value");
			var oFilter = new Filter("Desti", FilterOperator.Contains, sValue);
			var oBinding = oEvent.getParameter("itemsBinding");
			oBinding.filter([oFilter]);
		},
		onValueHelpDestinoClose: function (oEvent) {
			var oSelectedItem = oEvent.getParameter("selectedItem"),
				oInput = this._inputDestino;

			// if (!oSelectedItem) {
			oInput.resetProperty("value");
			// return;
			// }

			oInput.setValue(oSelectedItem.getDescription());
			let oDestino = this.getModel("viewModel").getProperty("/destino").filter(function (element) {
				return element.Desti === oSelectedItem.getDescription();
			})
			var oItems = this.getModel("viewModel").getData().items;
			for (var a = 0; a < oItems.length; a++) {
				if (oItems[a].Destino === oSelectedItem.getDescription()) {
					oItems[a].DestinoVs = oSelectedItem.getTitle();
					if (oDestino !== undefined && oDestino !== null && oDestino !== '') {
						oItems[a].TipoRecep = oDestino[0].TipoRecep;
					}
				};
			};
			var oModel = this.getView().getModel("viewModel");
			oModel.setProperty("/items", oItems);
		},
	intervaloSecundarioEnIntervaloPrimario:	function(fechaIniPrimario, fechaFinPrimario, fechaIniSecundario, fechaFinSecundario) {
	    // Convertir las fechas de string a objetos Date para el intervalo primario
	 
		let dateIniPrimario = parseInt(fechaIniPrimario.substring(0, 4));
		let dateFinPrimario = parseInt(fechaFinPrimario.substring(0, 4));
	    let dateIniSecundario = parseInt(fechaIniSecundario.substring(0, 4));
		let dateFinSecundario = parseInt(fechaFinSecundario.substring(0, 4));
    // Verificar si el intervalo secundario est� dentro del intervalo primario
    return dateIniSecundario >= dateIniPrimario && dateFinSecundario <= dateFinPrimario;
	},
		_onViewMatched: function (oEvent) {
			var that = this;
			var oViewModel = this.getView().getModel();
			var parametros = {
				'Kokrs': this.byId("inpSociedadCo").getValue().toUpperCase(),
				'Bukrs': this.byId("inpSociedadFi").getValue().toUpperCase()
			}
			return new Promise(function (resolve, reject) {
			oViewModel.setProperty("/busy", true);
			
				oViewModel.callFunction("/GetParametrosIniciales", {
				urlParameters: parametros,
				method: "GET",
				success: function (oData) {
				//	oViewModel.setProperty("/busy", false);
					if (oData.GetParametrosIniciales && oData.GetParametrosIniciales.PermitidoFechas) {
					var arrFechaIni = this.getView().byId("dpIper").getValue().split(".");
					var arrFechaFin = this.getView().byId("dpFper").getValue().split(".");
					var fechaIni =  arrFechaIni[2] +  arrFechaIni[1] + arrFechaIni[0];
					var fechaFin =  arrFechaFin[2] + arrFechaFin[1] + arrFechaFin[0];
					var validacion = that.intervaloSecundarioEnIntervaloPrimario(oData.GetParametrosIniciales.FechaInicio, oData.GetParametrosIniciales.FechaFin, fechaIni, fechaFin);
                    if (validacion == false)
					{
						that.showGeneralError({
								message: "La aplicación no se puede ejecutar en las fechas indicadas.",
								onClose: function () {
									that.onNavBack();
									reject(false);
								}.bind(that)
							});
					}
					resolve(true);
					} else {
						if (!oData.GetParametrosIniciales) {
							that.showGeneralError({
								message: "Hay un error al verificar los parámetros de la aplicación. Por favor consultar con el administrador del sistema.",
								onClose: function () {
									that.onNavBack();
									resolve(true);
								}.bind(that)
							});
						} else {
							that.showGeneralError({
								message: "La aplicación no se puede ejecutar en las fechas indicadas.",
								onClose: function () {
									that.onNavBack();
									reject(false);
								}.bind(that)
							});
						}
					}
				}.bind(that),
				error: function (oError) {
					that.showGeneralError({
						oDataError: oError
					
					});
					oViewModel.setProperty("/busy", false);
				}.bind(that)
			});
			});
		},
		_createViewModel: function () {
			var oModel = new JSONModel({
				SociedadFi: "",
				SociedadCo: "",
				busy: false,
				Ceco: "",
				Compa: "",
				personigrama: {},
				items: [],
				ceco: [],
				actividades: [],
				destino: [],
				newItem: {
					actividad: ""
				}
			});
			this.getView().setModel(oModel, "viewModel");

		},
		_getsociedadesFI: function () {
			var that = this;
			this.getOwnerComponent().getModel().read("/SociedadSet", {
				success: function (oData) {
					that.getOwnerComponent().setModel(oData.results, "SociedadFI");

				},
				error: function (oError) {
				
				}
			});

		},
		_getnewLine: function () {
			return {
				Ceco: "",
				CecoName: "",
				Actividad: "",
				ActividadName: "",
				Destino: "",
				DestinoVS: "None",
				mes1: "",
				mes2: "",
				mes3: "",
				mes4: "",
				mes5: "",
				mes6: "",
				mes7: "",
				mes8: "",
				mes9: "",
				mes10: "",
				mes11: "",
				mes12: ""
			};
		},
		_getActividades: function (sCeco, sYear, sIper, sFper, co) {
			var that = this;
			var aFilter = [];

			aFilter.push(new Filter("Kostl", FilterOperator.EQ, sCeco));
			aFilter.push(new Filter("Gjahr", FilterOperator.EQ, sYear));
			aFilter.push(new Filter("Iper", FilterOperator.EQ, sIper));
			aFilter.push(new Filter("Fper", FilterOperator.EQ, sFper));
			aFilter.push(new Filter("Kokrs", FilterOperator.EQ, co));


			return new Promise(function (resolve, reject) {
				that.getModel().read("/ActividadSet", {
					filters: aFilter,
					success: function (oData) {
						resolve(oData);
					},
					error: function (oError) {
						that.errorbackend(oError);
						reject(oError);
					}
				});
			});
		},
		_getDestinos: function (sCeco, sYear, sIper, sFper, co, fi) {
			var that = this;
			var aFilter = [];

			aFilter.push(new Filter("Kostl", FilterOperator.EQ, sCeco));
			aFilter.push(new Filter("Gjahr", FilterOperator.EQ, sYear));
			aFilter.push(new Filter("Iper", FilterOperator.EQ, sIper));
			aFilter.push(new Filter("Fper", FilterOperator.EQ, sFper));
			aFilter.push(new Filter("Comp", FilterOperator.EQ, fi));
			aFilter.push(new Filter("Kokrs", FilterOperator.EQ, co));
			return new Promise(function (resolve, reject) {
				that.getModel().read("/DestinoSet", {
					filters: aFilter,
					success: function (oData) {
						resolve(oData);
					},
					error: function (oError) {
						that.errorbackend(oError);
						reject(oError);
					}
				});
			});
		},
		_getPlantilla: function (sCeco, sYear, sIper, sFper, co, fi) {
			var that = this;
			var aFilter = [];
			aFilter.push(new Filter("Kostl", FilterOperator.EQ, sCeco));
			aFilter.push(new Filter("Gjahr", FilterOperator.EQ, sYear));
			aFilter.push(new Filter("Iper", FilterOperator.EQ, sIper));
			aFilter.push(new Filter("Fper", FilterOperator.EQ, sFper));
			aFilter.push(new Filter("Comp", FilterOperator.EQ, fi));
			aFilter.push(new Filter("Kokrs", FilterOperator.EQ, co));
			return new Promise(function (resolve, reject) {
				that.getModel().read("/PlantillaSet", {
					filters: aFilter,
					success: function (oData) {
						resolve(oData);
					},
					error: function (oError) {
						reject(oError);
					}
				});
			});
		},

		_getPersonigrama: function (sCeco, sYear, co, fi) {
			var that = this;
			var sPath = this.getModel().createKey("PersonigramaSet", {
				Gjahr: sYear,
				Kostl: sCeco,
				Kokrs: co,
				Bukrs: fi,
			});

			var oViewModel = this.getModel("viewModel");

			return new Promise(function (resolve, reject) {
				that.getModel().read("/" + sPath, {
					success: function (oData) {
						resolve(oData);
					},
					error: function (oError) {
						oViewModel.setProperty("/busy", false);
						that.showGeneralError({
							message: "No se encontró personigrama asociado."
						});
						reject(oError);
					}
				});
			});
		},

		_getData: function () {
			var oView = this.getView();
			var oViewModel = this.getModel("viewModel");
			var sCeco = oViewModel.getProperty("/Ceco").toUpperCase();
			var dIper = oViewModel.getProperty("/Iper");
			var sComp = oViewModel.getProperty("/Compa").toUpperCase();
			var aItems = oViewModel.getProperty("/items");
			var oPersonigrama = oViewModel.getProperty("/personigrama");
			var sYear = dIper.getFullYear().toString();
			var oObject = {
				//Kostl: Formatter.addLeftZeros(sCeco, 10),
				Kostl: sCeco,
				Gjahr: sYear,
				Items: []
			};
			for (var i = 0; i < aItems.length; i++) {
				var oItem = {
					Kostl: oObject.Kostl,
					Gjahr: sYear,
					Lstar: aItems[i].Actividad,
					Desti: aItems[i].Destino.toString(),
					Comp: oView.byId("inpSociedadFi").getValue().toUpperCase(),
					DestiDesc: aItems[i].DestinoVs,
					TipoRecep: aItems[i].TipoRecep
				};
				for (var j = 1; j <= 12; j++) {
					var fValue = parseFloat(aItems[i]['mes' + j], 10);
					if (isNaN(fValue)) {
						fValue = 0;
					}
					oItem['Mes' + j] = fValue.toString();
				}
				oObject.Items.push(oItem);
			}
			return oObject;
		},
		errorbackend: function (oError) {
			var oViewModel = this.getModel("viewModel");
			oViewModel.setProperty("/busy", false);
			var mensaje = JSON.parse(oError.responseText).error.message.value;
			sap.m.MessageBox.error(mensaje || "Ocurrió un error al procesar la solicitud. Por favor, intente nuevamente o contacte al administrador del sistema.");
		},

		_checkDatosPlantilla: function () {
			var oViewModel = this.getModel("viewModel");
			var aConsolidado = this.getModel("TotalesModel").getProperty("/List");
			let buttonSave = this.byId("buttonSave");
			var bError = false;
			for (var i = 1; i <= 12; i++) {
				var sValueState = "Success";
				if (Math.round(aConsolidado[0]["mes" + i]) !== parseFloat(aConsolidado[3]["mes" + i])) {
					bError = true;
				}
				if(Math.round(aConsolidado[0]["mes" + i]) >= parseFloat(aConsolidado[3]["mes" + i])){
					buttonSave.setEnabled(false);
				}else{
					buttonSave.setEnabled(true);
				}
			}
			var aList = oViewModel.getProperty("/items");
			for (var i = 0; i < aList.length; i++) {
				if (i !== (aList.length - 1)) {
					if (!aList[i].Destino) {
						aList[i].DestinoVS = "Error";
						bError = true;
					}
				}
			}
			oViewModel.refresh();
			return bError;
		},
		_getReal: function (e, r, o, y) {
			var i = this;
			var n = [];
			n.push(new Filter("Kostl", FilterOperator.EQ, e));
			n.push(new Filter("Iper", FilterOperator.EQ, o));
			n.push(new Filter("Fper", FilterOperator.EQ, y));
			n.push(new Filter("Lyear", FilterOperator.EQ, r));
			return new Promise(function (e, t) {
				i.getModel().read("/RealSet", {
					filters: n,
					success: function (t) {
						e(t)
					},
					error: function (e) {
						t(e)
					}
				})
			})
		},
		onChangeRecep: function (oEvt) {
			var selectedItem = oEvt.getParameter("selectedItem"),
				oModel = this.getModel("viewModel"),
				oText = selectedItem.getProperty("text"),
				aItem = oEvt.getSource().getBindingContext("viewModel").getObject();
			aItem.TipoRecep = oText;
		},
		onLiveChange: function (oEvt) {
			var oValue = oEvt.getParameter("value");

			if (oValue.substring(0, 1) === '-') {
				oEvt.getSource().setValueState("Error");
				oEvt.getSource().setValueState("No se admiten negativos");
			} else {
				oEvt.getSource().setValueState("Success");
			}
		},
		disableFields: function (oBoolean) {
			this.getModel("viewModel").setProperty("/editable", oBoolean);
		},

		validateCapMax: async function(aPlantilla,oPersonigrama){
			let messageError = false;
			for (let m of aPlantilla) {
				for (let i = 1; i <= 12; i++) {
					let mesPlantilla = parseFloat(m[`Mes${i}`]);
					let capMaxMes = parseFloat(oPersonigrama[`CapMaxMes${i}`]);
						if (mesPlantilla > capMaxMes) {
							messageError=true;
							return; 
						}
					}
				}
			if(messageError){
				MessageBox.error("Hay meses donde el porcentaje de capacidad es superior al 100%. No se permite continuar.");		
			}
		}
	});
});