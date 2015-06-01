function showAll()
{
	var nameOnly = new Office.Controls.Persona('nameOnlyRoot', 'NameOnly', true);
	Office.Controls.Utils.addEventListener(nameOnly.root, 'click', function (e) {
        return new Office.Controls.Persona('nameOnlyRoot', 'NameImage', true);
    });
	
	var personaCard = new Office.Controls.Persona('personaCardRoot', 'PersonaCard', true);
}