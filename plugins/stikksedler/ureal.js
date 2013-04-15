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

    // Utlånseddel: Felles uavhengig av språk
    reg_common: function (doc, user, library, excel) {
        if (user.kind === 'person') {
            excel.Cells(1, 1).Value = user.etternavn + ', ' + user.fornavn;
        }
        excel.Cells(7, 3).Value = doc.tittel;
        excel.Cells(8, 3).Value = doc.dokid;
    },

    // Utlånseddel på norsk
    reg_no: function (doc, user, library) {
        var excel = this.load_xls('plugins/stikksedler/ureal/reg_no.xls');
        this.reg_common(doc, user, library, excel);

        excel.Cells(2, 1).Value = " Utlånsdato : " + doc.utlaansdato;
        if (doc.forfvres !== doc.forfallsdato) {
            excel.Cells(3, 2).Value = "Lånefrist / Due : " + doc.forfallsdato;
            excel.Cells(4, 2).Value = "Ved reservasjoner kan documentet bli innkalt fra: " + doc.forfvres;
        } else {
            excel.Cells(3, 2).Value = "Lånefrist / Due : " + doc.forfvres;
        }

        // Er låner et bibliotek?
        if (user.kind === 'bibliotek') {
            excel.Cells(1, 1).Value  = "Fjernlån til " + user.navn;
            excel.Cells(32, 1).Value = user.ltid.substr(3,3);   // Venstre del av lib-nr.
            excel.Cells(32, 4).Value = user.ltid.substr(6);     // Høyre del av lib-nr.

        // Skal boka til et annet bibliotek innad i organisasjonen?
        } else if (user.beststed !== this.beststed) {
            excel.Cells(32, 1).Value = library.ltid.substr(3,3);    // Venstre del av lib-nr.
            excel.Cells(32, 4).Value = library.ltid.substr(6);      // Høyre del av lib-nr.
            excel.Cells(31, 1).Value = config.biblnavn[library.ltid];
        }

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

        //excel.Cells(7, 1).Value = "Tittel:"

        // Vise norsk logo
        //excel.ActiveSheet.Shapes("Picture 3").Visible = true;
        
        // Skjule engelsk logo
        //excel.ActiveSheet.Shapes("Picture 2").Visible = false;

        this.print_and_close();   
    },

    // Utlånseddel på engelsk
    reg_en: function (doc, user, library) {
        var excel = this.load_xls('plugins/stikksedler/ureal/reg_en.xls');
        this.reg_common(doc, user, library, excel);

        excel.Cells(2, 1).Value = " Loan date : " + doc.utlaansdato;
        if (doc.forfvres !== doc.forfallsdato) {
            excel.Cells(3, 2).Value = "Due : " + doc.forfallsdato;
            excel.Cells(4, 2).Value = "If required by another loaner, the document may be recalled from: " + doc.forfvres;
        } else {
            excel.Cells(3, 2).Value = "Due : " + doc.forfvres;
        }

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

        //excel.Cells(7, 1).Value = "Title:"

        // Skjule norsk logo
        //excel.ActiveSheet.Shapes("Picture 3").Visible = false;
        
        // Vise engelsk logo
        //excel.ActiveSheet.Shapes("Picture 2").Visible = true;

        this.print_and_close();
    },

    ret_en: function(doc, user, library) {
        var excel = this.load_xls('plugins/stikksedler/ureal/ret_en.xls');
        excel.Cells(7, 3).Value = doc.tittel;
        excel.Cells(8, 3).Value = doc.dokid;
        excel.Cells(2, 1).Value = "Return date: " + this.current_date();
        excel.Cells(8, 3).Value = doc.bestnr;

        // excel.Cells(1, 1).Value = "Returned with thanks "
        // excel.Cells(3, 2).Value = "Science Library - University of Oslo"
        // excel.Cells(7, 1).Value = "Title: "
        // excel.Cells(8, 1).Value = "Order no: "

        // excel.Cells(11, 1).Value = ""
        // excel.Cells(12, 1).Value = ""
        // excel.Cells(18, 2).Value = ""
        // excel.Cells(19, 2).Value = ""
        // excel.Cells(20, 2).Value = ""
        // excel.Cells(21, 2).Value = ""
        
        // //Skjule norsk logo
        // excel.ActiveSheet.Shapes("Picture 3").Visible = false;
        
        // // Vise engelsk logo
        // excel.ActiveSheet.Shapes("Picture 2").Visible = true;
        
        // // skjule verktøy-knapp for bibsys
        // excel.ActiveSheet.Shapes("Picture 1").Visible = false;
    },

    ret_no: function(doc, user, library) {
        var excel = this.load_xls('plugins/stikksedler/ureal/ret_no.xls');
        excel.Cells(7, 3).Value = doc.tittel;
        excel.Cells(8, 3).Value = doc.dokid;
        excel.Cells(2, 1).Value = "Returdato:" + this.current_date();

        // IRET ?
        if (doc.bestnr !== '') {
            //excel.Cells(1, 1).Value = 'Retur fra Realfagsbiblioteket';
            excel.Cells(31, 1).Value = lib.navn;
        }

        excel.Cells(32, 1).Value = library.ltid.substr(3,3);  //Venstre del av lib-nr.
        excel.Cells(32, 3).Value = library.ltid.substr(6);    //Høyre del av lib-nr.            

        // excel.Cells(1, 1).Value = "Retur fra UREAL"
        // excel.Cells(3, 2).Value = "Takk for lånet!"
        
        // excel.Cells(11, 1).Value = ""
        // excel.Cells(12, 1).Value = ""
        // excel.Cells(18, 2).Value = ""
        // excel.Cells(19, 2).Value = ""
        // excel.Cells(20, 2).Value = ""
        // excel.Cells(21, 2).Value = ""

        // // Vise norsk logo
        // excel.ActiveSheet.Shapes("Picture 3").Visible = true;
        
        // // Skjule engelsk logo
        // excel.ActiveSheet.Shapes("Picture 2").Visible = false;
        
        // // skjule verktøy-knapp for bibsys
        // excel.ActiveSheet.Shapes("Picture 1").Visible = false;
    }

});
