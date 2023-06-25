Option Explicit

Public Function GetCateringMenu() As String

'Declare Variables
Dim CateringMenu As String

'Create the Catering Menu
CateringMenu = "A catering service that provides high-quality and delicious food for events, parties, and special occasions. "
CateringMenu = CateringMenu & "Our extensive menu offers choices including: "
CateringMenu = CateringMenu & "Breakfast: "
CateringMenu = CateringMenu & "Breakfast Burritos, Pancakes, French Toast, Bagels, Scrambled Eggs, Breakfast Sandwiches "
CateringMenu = CateringMenu & "Lunch: "
CateringMenu = CateringMenu & "Sandwiches, Salads, Soup, Wraps, Burgers, Pizza, Burrito Bowls "
CateringMenu = CateringMenu & "Dinner: "
CateringMenu = CateringMenu & "Pasta, Steak, Seafood, Chicken, Vegetarian Dishes, Tacos, BBQ Dishes "
CateringMenu = CateringMenu & "Additional Options: "
CateringMenu = CateringMenu & "Assorted Appetizers, Desserts, Beverages "

'Return Result
GetCateringMenu = CateringMenu

End Function

Public Function GetCateringPackageOptions() As String

'Declare Variables
Dim CateringPackageOptions As String

'Create the Catering Package Options
CateringPackageOptions = "We offer multiple packages to choose from: "
CateringPackageOptions = CateringPackageOptions & "Standard: includes three entrees and two side dishes. "
CateringPackageOptions = CateringPackageOptions & "Deluxe: includes four entrees and three side dishes. "
CateringPackageOptions = CateringPackageOptions & "Premium: includes five entrees and four side dishes. "
CateringPackageOptions = CateringPackageOptions & "Custom: create your own menu with any combination of entrees and side dishes. "

'Return Result
GetCateringPackageOptions = CateringPackageOptions

End Function

Public Sub GetCateringServiceDetails()

'Declare Variables
Dim CateringServiceDetails As String

'Create the Catering Service Details
CateringServiceDetails = "Our catering services are perfect for any event, from small intimate gatherings to large corporate events. "
CateringServiceDetails = CateringServiceDetails & "We will work with you to customize the menu to meet your needs and budget. "
CateringServiceDetails = CateringServiceDetails & "We provide full setup and cleanup services for all events. "
CateringServiceDetails = CateringServiceDetails & "Our team of experienced staff will ensure that your event runs smoothly and is a success. "

'Display the Catering Service Details
MsgBox CateringServiceDetails

End Sub