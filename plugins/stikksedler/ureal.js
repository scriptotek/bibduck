$.extend($.bibduck.stikksedler, {

    // Utlånseddel
    reg: function(doc, user, library) {
        if (user.spraak === 'ENG') {
            this.reg_en(doc, user, library);
        } else {
            this.reg_no(doc, user, library);
        }
    },

    // Returseddel
    ret: function(doc, user, library) {
        if (library.navn === 'xxx') {
            // Retur til utlandet
            this.ret_en(doc, user, library);
        } else {
            this.ret_no(doc, user, library);
        }
    },

	format_date: function(dt, lang) {
		var fdato = dt.split('-');
		if (lang === 'ENG') {
			return fdato[2] + '. ' + month_names_en[fdato[1]-1] + ' ' + fdato[0];
		} else {
			return fdato[2] + '. ' + month_names[fdato[1]-1] + ' ' + fdato[0];
		}
	},

    // Utlånseddel: Felles uavhengig av språk
    template_replacements: function (doc, user, library, excel) {
        var cells = new Enumerator(excel.ActiveSheet.UsedRange.Cells),
            cell,
            libv = '',
            libh = '',
            navn = user.etternavn + ', ' + user.fornavn;
		
        if (user.kind === 'bibliotek') {
            libv = user.ltid.substr(3,3),
            libh = user.ltid.substr(6);
            navn = 'Fjernlån';  // til ' + user.navn;
        } else if (user.beststed !== this.beststed) {
            libv = library.ltid.substr(3,3);    // Venstre del av lib-nr.
            libh = library.ltid.substr(6);      // Høyre del av lib-nr.
            // excel.Cells(31, 1).Value = config.biblnavn[library.ltid];
        }
		
		if (doc.utlaansdato === undefined) doc.utlaansdato = this.current_date();
		if (doc.forfallsdato === undefined) doc.forfallsdato = this.current_date();
		if (doc.forfvres === undefined) doc.forfvres = this.current_date();

        for (; !cells.atEnd(); cells.moveNext()) {
            cell = cells.item();
            if (cell.Value !== undefined && cell.Value !== null) {
                cell.Value = cell.Value.replace('{{Navn}}', navn)
                                    .replace('{{Libnavn}}', library.navn)
                                    .replace('{{Tittel}}', doc.tittel)
                                    .replace('{{Dokid}}', doc.dokid)
                                    .replace('{{Utlånsdato}}', this.format_date(doc.utlaansdato, user.spraak))
                                    .replace('{{Forfallsdato}}', this.format_date(doc.forfallsdato, user.spraak))
                                    .replace('{{ForfallVedRes}}', this.format_date(doc.forfvres, user.spraak))
                                    .replace('{{LIBV}}', libv)
                                    .replace('{{LIBH}}', libh)
                                    .replace('{{Dato}}', this.format_date(this.current_date()))
                                    .replace('{{Bestnr}}', doc.bestnr);
            }
        }

        // Hvis forfallsdato ved reservasjon er lik ordinær forfallsdato:
        if (doc.forfvres === doc.forfallsdato) {
            excel.Cells(4, 2).Value = '';
        }
    },

    // Utlånseddel på norsk
    reg_no: function (doc, user, library) {
        var excel = this.load_xls('plugins/stikksedler/ureal/reg_no.xls');
        this.template_replacements(doc, user, library, excel);

        // Skal boka til et annet bibliotek innad i organisasjonen?
        // Hvis ikke fjernlån, skriv ut litt ekstra info om fornying:
        if (user.kind === 'person') {
            if (doc.purretype === 'E') {
                if (doc.utlstatus === 'UTL/RES') {
                    excel.Cells(11, 1).Value = "NB:";
                    excel.Cells(12, 1).Value = "Dette dokumentet kan ikke fornyes, da det er reservert for en annen låntaker.";
                } else {
                    excel.Cells(11, 1).Value = "Dette lånet kan du ikke fornye selv på BIBSYS Ask.";
                    excel.Cells(12, 1).Value = "Kom innom biblioteket hvis du ønsker å fornye dette lånet.";
                }
            } else {
                excel.Cells(11, 1).Value = "Dette lånet kan du fornye selv på BIBSYS Ask";
                excel.Cells(12, 1).Value = "hvis det ikke kommer reserveringer.";
            }
        }

        this.print_and_close();
    },

    // Utlånseddel på engelsk
    reg_en: function (doc, user, library) {
        var excel = this.load_xls('plugins/stikksedler/ureal/reg_en.xls');
        this.template_replacements(doc, user, library, excel);

        // Hvis ikke fjernlån, skriv ut litt ekstra info om fornying:
        if (user.kind === 'person') {
            if (doc.purretype === 'E') {
                if (doc.utlstatus === 'UTL/RES') {
                    excel.Cells(11, 1).Value = "Please note:";
                    excel.Cells(12, 1).Value = "This document can not be renewed as it has been reserved by someone else.";
                } else {
                    excel.Cells(11, 1).Value = "This document can not be renewed online at BIBSYS Ask.";
                    excel.Cells(12, 1).Value = "Please visit the library if you want to renew it.";
                }
            } else {
                excel.Cells(11, 1).Value = "Unless requested by someone else, this document can be";
                excel.Cells(12, 1).Value = "renewed online at BIBSYS Ask.";
            }
        }

        this.print_and_close();
    },

    ret_en: function(doc, user, library) {
        var excel = this.load_xls('plugins/stikksedler/ureal/ret_en.xls');
        this.template_replacements(doc, user, library, excel);
        this.print_and_close();
    },

    ret_no: function(doc, user, library) {
        var excel = this.load_xls('plugins/stikksedler/ureal/ret_no.xls');
        this.template_replacements(doc, user, library, excel);
        this.print_and_close();
    }

});
