<#
Create-Presentation.ps1

This PowerShell script generates an editable PowerPoint (.pptx)
presentation for the InventoryMVC project using PowerPoint COM automation.

Requirements:
- Microsoft PowerPoint must be installed on the machine where this script runs.
- Run from PowerShell with execution allowed: powershell -ExecutionPolicy Bypass -File .\Create-Presentation.ps1

Output:
- InventoryMVC_Presentation.pptx will be created in the same folder as this script (project tools folder).

Usage:
1. Open PowerShell in the workspace folder (E:\MCA\SEM_3\NET\WebApplication1\tools).
2. Run: powershell -ExecutionPolicy Bypass -File .\Create-Presentation.ps1

#>

try {
	$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
	$outputPath = Join-Path $scriptDir "..\InventoryMVC_Presentation.pptx" | Resolve-Path -Relative
	$outputFull = Join-Path $scriptDir "..\InventoryMVC_Presentation.pptx"

	Write-Host "Creating presentation at: $outputFull"

	$ppt = New-Object -ComObject PowerPoint.Application
	$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
	$presentation = $ppt.Presentations.Add()

	# Helper: add a slide with title and body text
	function Add-TitleBodySlide($titleText, $bodyText, $slideIndex, [bool]$addPlaceholder = $false) {
		$layout = $presentation.SlideMaster.CustomLayouts.Item(2) # Title and Content layout
		$slide = $presentation.Slides.AddSlide($slideIndex, $layout)
		if ($slide.Shapes.Count -gt 0) {
			if ($null -ne $slide.Shapes.Title) { $slide.Shapes.Title.TextFrame.TextRange.Text = $titleText }
			# Find first placeholder for body
			foreach ($shape in $slide.Shapes) {
				if ($shape.Type -eq 14 -or $shape.PlaceholderFormat -ne $null) {
					try { $shape.TextFrame.TextRange.Text = $bodyText; break } catch { }
				}
			}
		}
		if ($addPlaceholder) {
			# Add a rectangle placeholder for screenshot
			$left = 50; $top = 180; $width = 560; $height = 300
			$rect = $slide.Shapes.AddShape(1, $left, $top, $width, $height) # msoShapeRectangle = 1
			$rect.Fill.Transparency = 0.9
			$rect.Line.ForeColor.RGB = 0x999999
			$rect.TextFrame.TextRange.Text = "Screenshot placeholder - add image here"
			$rect.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # center
		}
		return $slide
	}

	$idx = 1

	# Slide 1: Title
	$layoutTitle = $presentation.SlideMaster.CustomLayouts.Item(1)
	$slide = $presentation.Slides.AddSlide($idx, $layoutTitle)
	$slide.Shapes.Title.TextFrame.TextRange.Text = "InventoryMVC - Product Inventory Management"
	$slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "ASP.NET MVC 5, C#, EF6 | Visual Studio - .NET Framework 4.7.2`nYour Name - Date"
	$idx++

	# Slide 2: Problem Statement
	Add-TitleBodySlide -titleText "Problem Statement" -bodyText "- Products managed in spreadsheets: error-prone, no validation, no concurrency`n- Need a simple web app to centralize product data, enforce validation, and enable CRUD operations." -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 3: Project Goals & Requirements
	Add-TitleBodySlide -titleText "Project Goals & Requirements" -bodyText "- ASP.NET MVC 5, C#, Entity Framework 6 (Code First), SQL Server LocalDB`n- Razor views, Bootstrap UI, full CRUD, server-side validation" -slideIndex $idx | Out-Null
	$idx++

	# Slide 4: Solution Overview (Architecture)
	Add-TitleBodySlide -titleText "Solution Overview (Architecture)" -bodyText "- MVC pattern: Controllers, Views, Models`n- EF6 Code First maps models to LocalDB`n- Bootstrap for responsive UI" -slideIndex $idx -addPlaceholder $true | Out-Null
	$idx++

	# Slide 5: Data Model (Product)
	$body = "Product model fields and validation:`n- ProductId (int, PK) `n- Name (string, required, max 100)`n- Description (string, optional)`n- Price (decimal, required, >=0)`n- Quantity (int, required, >=0)`n- Category (string, optional, max 50)"
	Add-TitleBodySlide -titleText "Data Model (Product)" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 6: Persistence and Connection
	$body = "InventoryContext : DbContext with DbSet<Product>`nConnection string in Web.config: (localdb)\\MSSQLLocalDB;Initial Catalog=InventoryMVCDB;Integrated Security=True`nUse Code First Migrations to create/update DB"
	Add-TitleBodySlide -titleText "Persistence & Connection" -bodyText $body -slideIndex $idx | Out-Null
	$idx++

	# Slide 7: Packages & Setup
	$body = "Required NuGet packages:`n- EntityFramework (6.x) - Install-Package EntityFramework -Version 6.4.4`n- Microsoft.AspNet.Mvc, Razor, jQuery, bootstrap (already present in project)" 
	Add-TitleBodySlide -titleText "Packages & Setup" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 8: Controller (ProductsController)
	$body = "Controller actions: Index, Details, Create (GET/POST), Edit (GET/POST), Delete (GET) + DeleteConfirmed (POST)`n- Uses [HttpPost], [ValidateAntiForgeryToken], ModelState.IsValid, EF operations" 
	Add-TitleBodySlide -titleText "Controller (ProductsController)" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 9: Views & UI
	$body = "Views: Index, Details, Create, Edit, Delete (`n) - Bootstrap table, forms with validation messages, anti-forgery tokens. `n- Strongly-typed Razor views" 
	Add-TitleBodySlide -titleText "Views & UI" -bodyText $body -slideIndex $idx -addPlaceholder $true | Out-Null
	$idx++

	# Slide 10: Demo Steps
	$body = "Demo steps:`n1. Start app (F5) - Products list opens`n2. Create product (show validation) `n3. View, Edit, Delete product" 
	Add-TitleBodySlide -titleText "Demo (Steps)" -bodyText $body -slideIndex $idx -addPlaceholder $true | Out-Null
	$idx++

	# Slide 11: Errors Encountered & Solutions
	$body = "Example issues and fixes:`n- Missing EF -> Install-Package EntityFramework 6.4.4`n- View not found -> Add Views/Products/Index.cshtml`n- Global.asax parse error -> Rebuild, ensure bin contains assembly, clear temp ASP.NET files" 
	Add-TitleBodySlide -titleText "Errors Encountered & Solutions" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 12: Validation & Security
	$body = "- Server-side validation via data annotations`n- Client-side with jquery.validate + unobtrusive`n- Anti-forgery tokens for POSTs`n- Use POST for changes and confirm delete" 
	Add-TitleBodySlide -titleText "Validation & Security" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 13: Project Structure & Key Files
	$body = "Key files:`n- Models/Product.cs, InventoryContext.cs`n- Controllers/ProductsController.cs`n- Views/Products/*.cshtml`n- App_Start/RouteConfig.cs, Web.config" 
	Add-TitleBodySlide -titleText "Project Structure & Key Files" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 14: Build & Run (step-by-step)
	$body = "1. Install EF: Install-Package EntityFramework -Version 6.4.4`n2. (Optional) Migrations: Enable-Migrations; Add-Migration InitialCreate; Update-Database`n3. Build solution, Run (F5)" 
	Add-TitleBodySlide -titleText "Build & Run" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 15: Test Cases to Show
	$body = "Manual tests:`n- Missing Name -> validation`n- Negative Price -> validation`n- Create, Edit, Delete flow`n- Verify DB table in SQL Server Object Explorer" 
	Add-TitleBodySlide -titleText "Test Cases to Show" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 16: Future Improvements
	$body = "- Add search, paging, and filters`n- Authentication & authorization`n- Soft-delete & audit logs`n- Export reports (CSV/Excel)" 
	Add-TitleBodySlide -titleText "Future Improvements" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 17: Appendix (Commands & References)
	$body = "Commands:`n- Install-Package EntityFramework -Version 6.4.4`n- Enable-Migrations; Add-Migration InitialCreate; Update-Database`nReferences: EF6 docs, ASP.NET MVC 5 docs, Bootstrap docs" 
	Add-TitleBodySlide -titleText "Appendix: Commands & References" -bodyText $body -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Slide 18: Q&A
	Add-TitleBodySlide -titleText "Q&A" -bodyText "Questions? Contact / GitHub repo: (add link)" -slideIndex $idx -addPlaceholder $false | Out-Null
	$idx++

	# Save presentation
	$presentation.SaveAs((Resolve-Path $outputFull).Path)
	Write-Host "Presentation saved: $outputFull"

	# Close
	$presentation.Close()
	$ppt.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null

} catch {
	Write-Error "Error creating presentation: $_"
	if ($ppt -ne $null) {
		try { $ppt.Quit() } catch { }
	}
	exit 1
}

Write-Host "Done. Open the generated file and edit content/screenshots as needed."
