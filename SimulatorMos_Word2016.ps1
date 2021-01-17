$tableau_titres = ("Tâche 1","Tâche 2","Tâche 3","Tâche 4")
$tableau_details = ('Au bas de la première page, triez le tableau en fonction du champ "Produit" par ordre croissant.',"details tâche 2","details tâche 3","details tâche 4")
$global:compteur_taches = 0
$max_taches = 4
 Function ClearAndClose()
 {
    $Timer.Stop(); 
    $Form.Close(); 
    $Form.Dispose();
    $Timer.Dispose();
    $Script:CountDown=5
 }

Function Terminer_Click()
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $result = [System.Windows.Forms.MessageBox]::Show("êtes-vous sûre de vouloire quitter l'examen ?" , "Quitter ?" , 3)
    if ($result -eq 'Yes') {
    ClearAndClose
    }
}
Function Suivant_Click()
{
    $process_word = get-process *word*
    if($process_word -ne ""){
        stop-process $process_word.Id
    }
    $global:compteur_taches++
    if ($global:compteur_taches -ge $max_taches)
    {
        [System.Windows.Forms.MessageBox]::Show("Vous avez terminé." , "Terminé" , 1)
        ClearAndClose

    }
    $Titre_tache.Text = $tableau_titres[$compteur_taches]
    $details_tache.Text = $tableau_details[$compteur_taches]
    Start-Process ((Resolve-Path "./MOS_Word2016_Projet1.docx").Path)
}

 Function Timer_Tick()
 {

    $tempsRestant.Text = "Tems restant : $Script:CountDown"
         --$Script:CountDown
         if ($Script:CountDown -lt 0)
         {
            ClearAndClose
         }
 }



Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = '1000,300'
$Form.Text = "Simulateur MOS Word 2016"

$tempsRestant = New-Object System.Windows.Forms.Label
$tempsRestant.Location = New-Object System.Drawing.Point(700,2)
$tempsRestant.Text = "Temps restant : $Script:CountDown"
$tempsRestant.Size = New-Object System.Drawing.Size(180,23)
$tempsRestant.Font = new-object System.Drawing.Font('Ariel',10,[System.Drawing.FontStyle]::Bold)
$Form.controls.Add($tempsRestant)

$Terminer = New-Object System.Windows.Forms.Button
$Terminer.Location = New-Object System.Drawing.Point(920,2)
$Terminer.Width = 60
$Terminer.Height = 20
$Terminer.Text = "Terminer"
$Form.controls.Add($Terminer)

$Suivant = New-Object System.Windows.Forms.Button
$Suivant.Location = New-Object System.Drawing.Point(920,260)
$Suivant.Width = 80
$Suivant.Height = 40
$Suivant.Text = "Suivant"
$Form.controls.Add($Suivant)

$Titre_tache = New-Object System.Windows.Forms.Label
$Titre_tache.Location = New-Object System.Drawing.Point(30,30)
$Titre_tache.Text = $tableau_titres[0]
$Titre_tache.Font = new-object System.Drawing.Font('Ariel',14,[System.Drawing.FontStyle]::Bold)
$Form.controls.Add($Titre_tache)

$details_tache = New-Object System.Windows.Forms.Label
$details_tache.Location = New-Object System.Drawing.Point(30,70)
$details_tache.Text = $tableau_details[$compteur_taches]
$details_tache.Width = "800"
$Form.controls.Add($details_tache)


$Suivant.Add_Click({Suivant_Click})
$Terminer.Add_Click({Terminer_Click})

$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = 1000
$Script:CountDown = 3000

$Timer.Add_Tick({ Timer_Tick})


$Timer.Start()
$Form.ShowDialog()