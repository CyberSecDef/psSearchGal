Begin{
	clear;
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	
	Class UI{
		static [UI]$UI = $null;
		
		$form = (iex "New-Object System.Windows.Forms.Form");
		
		[UI] static Get(  ){
			if( [UI]::UI -eq $null){
				[UI]::UI = [UI]::new( );
			}

			return [UI]::UI
		}
		
		showBusinessCard($user){
			$bcardForm = (iex "New-Object System.Windows.Forms.Form");
			$bcardForm.Size = New-Object System.Drawing.Size(800,600)
			$bcardForm.StartPosition = 'CenterScreen'
			$bcardForm.FormBorderStyle = 'FixedDialog'
			$bcardForm.Text = $user.'Display Name'
			
			$bcardHTML = new-object System.Windows.Forms.WebBrowser
			$bcardHTML.name = 'bcard'
			$bcardHTML.size = New-Object System.Drawing.Size(800,600)
			$bcardHTML.location = New-Object System.Drawing.Point(0,0)
			$bcardHTML.DocumentText = @"
<html>
<head>
<style>
body{
	font-family: arial;
	font-size:10pt;
}
li{
	font-size:8pt;
}
table{
	width:100%;
}
td{
	width:50%;
	vertical-align:top;
}
</style>
</head>
<body>
<h1>$($user.'Employee Type') $($user.'First Name') $($user.'Last Name') ( $($user.'Samaccountname') )</h1>
<h2>Title : $($user.'Title')</h2>
<h3>EDIPI : $($user.'EDIPI')</h3>
<h4>Company : $($user.'Company')</h4>
<h4>Email : $($user.'Email')</h4>
<table>
<tr><td>
<ul>
	<li>Phone : $($user.'Phone')</li>
	<li>Mobile : $($user.'Mobile')</li>
	<li>Home : $($user.'Home Phone')</li>
</ul>
</td><td>
<ul>

<li>Division : $($user.'Division')</li>
<li>Base : $($user.'Base')</li>
<li>Department : $($user.'Department')</li>
<li>Building : $($user.'Building')</li>
<li>Room : $($user.'Room')</li>
</ul>
</td>
</tr>
</table>
Street Address:<br>
$($user.'Street Address')
</body>
</html>
"@			
			
			
			$bcardForm.controls.add($bcardHTML)
			$bcardForm.ShowDialog()
		}
		
		UI(){
			
			$this.form.Text = 'GAL Search Form'
			$this.form.Size = New-Object System.Drawing.Size(300,170)
			$this.form.StartPosition = 'CenterScreen'
			$this.form.FormBorderStyle = 'FixedDialog'

			$ComboBox = New-Object System.Windows.Forms.ComboBox
			$ComboBox.name = 'searchField'	
			$ComboBox.Location = New-Object System.Drawing.Point(10,40)
			$ComboBox.Size = New-Object System.Drawing.Size(100,20)
			$ComboBox.Height = 80

			@("First Name","Last Name","Base","Employee Type","EDIPI","Division","Department","Building","Room","Phone","Title","Email") | sort | % {
				[void] $ComboBox.Items.Add($_)
			}
			$ComboBox.selectedIndex = 0
			$this.form.Controls.Add($ComboBox)
			
			$valText = New-Object System.Windows.Forms.TextBox
			$valText.Name = 'searchTerm'
			$valText.text = ''
			$valText.Location = New-Object System.Drawing.Point(120,40)
			$valText.Size = New-Object System.Drawing.Size(100,20)
			$this.form.Controls.Add($valText)

			$valBtn = New-Object System.Windows.Forms.Button
			$valBtn.Location = New-Object System.Drawing.Point(230,40)
			$valBtn.Size = New-Object System.Drawing.Size(50,18)
			$valBtn.Text = 'Add'
			$valBtn.add_click({
				if( [UI]::Get().form.controls['searchTerm'].text.toLower().replace(' ','') -ne ''){
					[UI]::Get().form.Controls['query'].Text += [UI]::Get().form.controls['searchField'].text.toLower().replace(' ','') 
					[UI]::Get().form.Controls['query'].Text += '='
					[UI]::Get().form.Controls['query'].Text += [UI]::Get().form.controls['searchTerm'].text.toLower().replace(' ','')
					[UI]::Get().form.Controls['query'].Text += ' '
				}
				[UI]::Get().form.controls['searchTerm'].text = ''
				[UI]::Get().form.controls['searchField'].selectedIndex = 0
			})
			
			$this.form.Controls.Add($valBtn)
			
			$OKButton = New-Object System.Windows.Forms.Button
			$OKButton.Location = New-Object System.Drawing.Point(75,100)
			$OKButton.Size = New-Object System.Drawing.Size(50,23)
			$OKButton.Text = 'OK'
			$OKButton.DialogResult = (iex "[System.Windows.Forms.DialogResult]::OK")
			$this.form.AcceptButton = $OKButton
			$this.form.Controls.Add($OKButton)

			$CancelButton = New-Object System.Windows.Forms.Button
			$CancelButton.Location = New-Object System.Drawing.Point(125,100)
			$CancelButton.Size = New-Object System.Drawing.Size(50,23)
			$CancelButton.Text = 'Cancel'
			$CancelButton.DialogResult = (iex "[System.Windows.Forms.DialogResult]::Cancel")
			$this.form.CancelButton = $CancelButton
			$this.form.Controls.Add($CancelButton)
			
			$aboutButton = New-Object System.Windows.Forms.Button
			$aboutButton.Location = New-Object System.Drawing.Point(175,100)
			$aboutButton.Size = New-Object System.Drawing.Size(50,23)
			$aboutButton.Text = '?'
			$aboutButton.add_click({ 
				iex @"
					[System.Windows.Forms.Messagebox]::Show( 
					"Welcome to the NMCI Gal Search Wizard.`n`nPlease enter as much information as you have about the user being searched.`n`nThe fields will use wildcard filters, so partial names or search terms are valid.",
					"Welcome to the GAL Search Wizard"
				);
"@
			})
			$this.form.Controls.Add($aboutButton)
			
			$label = New-Object System.Windows.Forms.Label
			$label.Location = New-Object System.Drawing.Point(10,10)
			$label.Size = New-Object System.Drawing.Size(280,20)
			$label.Text = 'Please enter the search query below:'
			$this.form.Controls.Add($label)

			$textBox = New-Object System.Windows.Forms.TextBox
			$textBox.Name = 'query'
			$textBox.text = ''
			$textbox.enabled = $false
			$textBox.Location = New-Object System.Drawing.Point(10,70)
			$textBox.Size = New-Object System.Drawing.Size(260,20)
			$this.form.Controls.Add($textBox)
			
			$this.form.Topmost = $true
			$this.form.select();
			$this.form.Controls['query'].select()
		}
	}
	
	class Searcher{
		$objects = @()
		$ADSearcher = $null
		
		Searcher(){
			$this.objects = @()
			$this.ADSearcher = (iex "New-Object System.DirectoryServices.DirectorySearcher( ( [ADSI]'' ) )")
		}
		
		Search($query){
			$query = $query.toLower()
				
			$filter = "(&(objectCategory=user)(&"

			$query.split(' ') | % {
				$term = $_.trim()
				if($term.indexOf('=') -gt 0){
					$key, $val = $term.split('=')
					switch($key){
						"lastname" {
							$filter += "(sn=$($val)*)"
							break;
						}
						"firstname" {
							$filter += "(givenname=$($val)*)"
							break;
						}
						
						"base" {
							$filter += "(base=$($val))"
							break;
						}
						"type" {
							$filter += "(employeetype=$($val))"
							break;
						}
						
						"edipi" {
							$filter += "(edipi=*$($val)*)"
							break;
						}
						"division" {
							$filter += "(division=*$($val)*)"
							break;
						}
						"department" {
							$filter += "(department=$($val)*)"
							break;
						}
						"building" {
							$filter += "(building=*$($val)*)"
							break;
						}
						"room" {
							$filter += "(roomnumber=*$($val)*)"
							break;
						}
						"phone" {
							$filter += "(|"
							$filter += "(telephonenumber=*$($val)*)"
							$filter += "(mobile=*$($val)*)"
							$filter += "(homephone=*$($val)*)"
							$filter += ")"
							break;
						}
						"title" {
							$filter += "(title=*$($val)*)"
							break;
						}
						"email" {
							$filter += "(mail=*$($val)*)"
							break;
						}
					}
				}
			}
			$filter += "))"
			
			$this.ADSearcher.filter = $filter
			
			$adResults = $this.ADSearcher.findall()
			$adResults | sort { $_.properties.sn} | sort { $_.properties.givenname} | % {
				$user = $_
				$props = $user.properties
				$Object = [psCustomObject]@{
					"First Name" = $props.givenname | out-string
					"Last Name" = $props.sn | out-string
					"Display Name" = $props.displayname | out-string
					Title = $props.title | out-string
					Description = $props.description | out-string
					Company = $props.compaby | out-string
					"Employee Type" = $props.employeetype | out-string
					Samaccountname = $props.samaccountname | out-string
					EDIPI = $props.edipi | out-string
					Email  = $props.mail | out-string
					Phone = $props.telephonenumber | out-string
					Mobile = $props.mobile | out-string
					"Home Phone" = $props.homephone | out-string
					Base = $props.base | out-string
					Division = $props.division | out-string
					Department = $props.department | out-string
					Building = $props.building  | out-string
					Room = $props.roomnumber  | out-string
					"Street Address" = $props.streetaddress | out-string				
				}
				$this.objects += $Object
			}
		}
	}	
}

Process{
	$result = [UI]::Get().form.ShowDialog()
	
	if ($result -eq [System.Windows.Forms.DialogResult]::OK){
		$searcher = [Searcher]::new()
		if([UI]::Get().form.Controls['query'].Text.trim() -ne ''){		
			$searcher.Search( [UI]::Get().form.Controls['query'].Text )
			if($searcher.objects.count -gt 0){
				$user = $searcher.objects | out-gridview -Title 'User(s)' -outputMode Single 
				if($user -ne $null){
					[UI]::Get().showBusinessCard( $user )
				}
			}else{
				[System.Windows.Forms.Messagebox]::Show("No Results Found","GAL Search Wizard");
			}
		}else{
			[System.Windows.Forms.Messagebox]::Show("Invalid Search Query Entered","GAL Search Wizard");
		}
	}
}

End{
	[UI]::UI = $null
}
