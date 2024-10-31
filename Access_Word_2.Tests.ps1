# Access_Word_2.Tests.ps1

# Import the script to be tested
. "$PSScriptRoot\Access_Word_2.ps1"

Describe "WordDocument Class Tests" {
    BeforeAll {
        # Initialize variables
        $DocFileName = "TestDoc.docx"
        $ScriptRoot = "C:\Test"
        $DocFilePath = $ScriptRoot

        # Create a mock WordDocument instance
        $wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)
    }

    Context "Custom Properties" {
        It "should create a custom property" {
            $wordDoc.Create_Property("TestProp", "TestValue")
            $propValue = $wordDoc.Read_Property("TestProp")
            $propValue | Should -Be "TestValue"
        }

        It "should read a custom property" {
            $propValue = $wordDoc.Read_Property("TestProp")
            $propValue | Should -Be "TestValue"
        }

        It "should update a custom property" {
            $wordDoc.Update_Property("TestProp", "UpdatedValue")
            $propValue = $wordDoc.Read_Property("TestProp")
            $propValue | Should -Be "UpdatedValue"
        }

        It "should delete a custom property" {
            $wordDoc.Delete_Property("TestProp")
            $propValue = $wordDoc.Read_Property("TestProp")
            $propValue | Should -BeNullOrEmpty
        }
    }

    Context "Document Properties" {
        It "should get document properties" {
            $properties = $wordDoc.Get_Properties("Both")
            $properties | Should -Not -BeNullOrEmpty
        }
    }

    Context "Environment Checks" {
        It "should check PC environment" {
            $envInfo = $wordDoc.Check_PC_Env()
            $envInfo["PCName"] | Should -Be $env:COMPUTERNAME
        }

        It "should check Word library" {
            $wordDoc.Check_Word_Library()
            $libraryPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
            Test-Path $libraryPath | Should -Be $true
        }
    }

    Context "Open and Close Document" {
        It "should open a document" {
            $wordDoc.Open_Document()
            $wordDoc.Document.FullName | Should -Not -BeNullOrEmpty
        }

        It "should close a document" {
            $wordDoc.Close_Document()
            $wordDoc.Document | Should -BeNull
        }
    }

    Context "Custom Properties with Mocks" {
        It "should create a custom property with mocks" {
            Mock -CommandName Set-Content

            $wordDoc.Create_Property("NewProp", "NewValue")

            Assert-MockCalled -CommandName Set-Content -Exactly 1 -Scope It
        }

        It "should read a custom property with mocks" {
            Mock -CommandName Set-Content

            $result = $wordDoc.Read_Property("NewProp")
            $result | Should -BeNullOrEmpty
        }

        It "should update a custom property with mocks" {
            Mock -CommandName Set-Content

            $wordDoc.Update_Property("NewProp", "UpdatedValue")

            Assert-MockCalled -CommandName Set-Content -Exactly 1 -Scope It
        }

        It "should delete a custom property with mocks" {
            Mock -CommandName Set-Content

            $wordDoc.Delete_Property("NewProp")

            Assert-MockCalled -CommandName Set-Content -Exactly 1 -Scope It
        }
    }

    Context "Word Processes" {
        It "should close Word processes" {
            Mock -CommandName Get-Process -MockWith { return @([PSCustomObject]@{ Id = 1234 }) }
            Mock -CommandName Stop-Process

            $wordDoc.Close_Word_Processes()

            Assert-MockCalled -CommandName Stop-Process -Exactly 1 -Scope It
        }

        It "should ensure Word is closed" {
            Mock -CommandName Get-Process -MockWith { return @([PSCustomObject]@{ Id = 1234 }) }
            Mock -CommandName Stop-Process

            $wordDoc.Ensure_Word_Closed()

            Assert-MockCalled -CommandName Stop-Process -Exactly 1 -Scope It
        }
    }
}