$CurrentDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$PasswordFile = Join-Path $CurrentDir "Password.txt"
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogFile = Join-Path $CurrentDir ($ScriptName + ".log")

function Log($message) {
    $now = Get-Date
    $log = $now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "`t"
    $log += $message

    Write-Output $log | Out-File -FilePath $LogFile -Encoding UTF8 -append

    return $log
}

Set-ExecutionPolicy -Scope Process RemoteSigned -Force
Log(Get-ExecutionPolicy)

$M365Credential = Get-Credential
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credential $M365Credential | Out-Null

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null

$SiteUrl = "https://tenant.sharepoint.com/sites/nic/"
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$SPCredential = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
$Context.Credentials = $SPCredential
$Context.ExecuteQuery()
Log("ClientContext Executed")

$ListName = "出社/在宅 状況"
$List = $Context.Web.Lists.GetByTitle($ListName)
$Context.Load($List)
$Context.ExecuteQuery()
Log("ListName GetByTitle Executed")

$IsValidInitialDate = $False
$IsValidLastDate = $False
$IsValidDuration = $False

do {
  if($IsValidInitialDate -eq $False) {
    $InputInitialDate = Read-Host "InitialDate? (M/d)"
    try {
      $InitialDate = Get-Date -Date $InputInitialDate

      $IsValidInitialDate = $True
      $TargetDate = $InitialDate
    }
    catch {
      continue
    }
  }

  if($IsValidLastDate -eq $False) {
    $InputLastDate = Read-Host "LastDate? (M/d)"

    try {
      $LastDate = Get-Date -Date $InputLastDate
      $IsValidLastDate = $True
    }
    catch {
      continue
    }
  }

  if($InitialDate -gt $LastDate) {
    $IsValidInitialDate = $False
    $IsValidLastDate = $False
  }
  else {
    $IsValidDuration = $True
  }

} while($IsValidInitialDate -eq $False -or $IsValidLastDate -eq $False -or $IsValidDuration -eq $False)

$Fields = $List.Fields
$Context.Load($Fields)
$Context.ExecuteQuery()
Log("Fields Load Executed")

while ($TargetDate -le $LastDate) {
    $FieldName = $TargetDate.ToString("M/d (ddd)")
    $Field = $List.Fields | Where-Object { $_.Title -eq $FieldName }
    if($Null -ne $Field) {
        $Field.DeleteObject()
        $Context.ExecuteQuery()
    }

    $TargetDate = $TargetDate.AddDays(1)
}

$TargetDate = $InitialDate

while ($TargetDate -le $LastDate) {
    $FieldID = New-Guid
    $Name = "Day" + $TargetDate.ToString("yyyyMMdd")
    $DisplayName = $TargetDate.ToString("M/d (ddd)")
    $Description = ""
    $IsRequired = $FALSE
    $EnforceUniqueValues = $FALSE
    $MaxLength = 255

    $FieldSchema = @"
        <Field Type='Choice' ID='{$FieldID}' Name='$Name' StaticName='$Name' DisplayName='$DisplayName' Description='$Description' Required='$IsRequired' EnforceUniqueValues='$EnforceUniqueValues' MaxLength='$MaxLength'>
            <CHOICES>
                <CHOICE>在宅（勤務）</CHOICE>
                <CHOICE>在宅（待機）</CHOICE>
                <CHOICE>出社</CHOICE>
                <CHOICE>外出</CHOICE>
                <CHOICE>休み</CHOICE>
                <CHOICE>有給休暇</CHOICE>
                <CHOICE>有給休暇（午前）</CHOICE>
                <CHOICE>有給休暇（午後）</CHOICE>
            </CHOICES>
        <Default>在宅（勤務）</Default>
        </Field>
"@

    $List.Fields.AddFieldAsXml($FieldSchema, $True, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
    $Context.ExecuteQuery()

    $TargetDate = $TargetDate.AddDays(1)
}

$TargetDate = $InitialDate

$JsonFormatSaturday = @"
{
  "`$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "style": {
    "box-sizing": "border-box",
    "padding": "0 2px"
  },
  "attributes": {
    "class": {
      "operator": ":",
      "operands": [
        {
          "operator": "==",
          "operands": [
            {
              "operator": "toLowerCase",
              "operands": [
                "@currentField"
              ]
            },
            {
              "operator": "toLowerCase",
              "operands": [
                "在宅（勤務）"
              ]
            }
          ]
        },
        "sp-css-backgroundColor-blueBackground37",
        {
          "operator": ":",
          "operands": [
            {
              "operator": "==",
              "operands": [
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "@currentField"
                  ]
                },
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "在宅（待機）"
                  ]
                }
              ]
            },
            "sp-css-backgroundColor-blueBackground27",
            {
              "operator": ":",
              "operands": [
                {
                  "operator": "==",
                  "operands": [
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "@currentField"
                      ]
                    },
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "出社"
                      ]
                    }
                  ]
                },
                "sp-css-backgroundColor-blueBackground17",
                {
                  "operator": ":",
                  "operands": [
                    {
                      "operator": "==",
                      "operands": [
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "@currentField"
                          ]
                        },
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "外出"
                          ]
                        }
                      ]
                    },
                    "sp-css-backgroundColor-blueBackground07",
                    {
                      "operator": ":",
                      "operands": [
                        {
                          "operator": "==",
                          "operands": [
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "@currentField"
                              ]
                            },
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "休み"
                              ]
                            }
                          ]
                        },
                        "sp-css-backgroundColor-successBackground50",
                        {
                          "operator": ":",
                          "operands": [
                            {
                              "operator": "==",
                              "operands": [
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "@currentField"
                                  ]
                                },
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "有給休暇"
                                  ]
                                }
                              ]
                            },
                            "sp-css-backgroundColor-successBackground40",
                            {
                              "operator": ":",
                              "operands": [
                                {
                                  "operator": "==",
                                  "operands": [
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "@currentField"
                                      ]
                                    },
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "有給休暇（午前）"
                                      ]
                                    }
                                  ]
                                },
                                "sp-css-backgroundColor-successBackground30",
                                {
                                  "operator": ":",
                                  "operands": [
                                    {
                                      "operator": "==",
                                      "operands": [
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "@currentField"
                                          ]
                                        },
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "有給休暇（午後）"
                                          ]
                                        }
                                      ]
                                    },
                                    "sp-css-backgroundColor-successBackground",
                                    {
                                      "operator": ":",
                                      "operands": [
                                        {
                                          "operator": "==",
                                          "operands": [
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "@currentField"
                                              ]
                                            },
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                ""
                                              ]
                                            }
                                          ]
                                        },
                                        "sp-css-backgroundColor-blueBackground07",
                                        ""
                                      ]
                                    }
                                  ]
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    }
  },
  "children": [
    {
      "elmType": "span",
      "attributes": {
        "iconName": {
          "operator": ":",
          "operands": [
            {
              "operator": "==",
              "operands": [
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "@currentField"
                  ]
                },
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "在宅（勤務）"
                  ]
                }
              ]
            },
            "",
            {
              "operator": ":",
              "operands": [
                {
                  "operator": "==",
                  "operands": [
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "@currentField"
                      ]
                    },
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "在宅（待機）"
                      ]
                    }
                  ]
                },
                "",
                {
                  "operator": ":",
                  "operands": [
                    {
                      "operator": "==",
                      "operands": [
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "@currentField"
                          ]
                        },
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "出社"
                          ]
                        }
                      ]
                    },
                    "",
                    {
                      "operator": ":",
                      "operands": [
                        {
                          "operator": "==",
                          "operands": [
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "@currentField"
                              ]
                            },
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "外出"
                              ]
                            }
                          ]
                        },
                        "",
                        {
                          "operator": ":",
                          "operands": [
                            {
                              "operator": "==",
                              "operands": [
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "@currentField"
                                  ]
                                },
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "休み"
                                  ]
                                }
                              ]
                            },
                            "",
                            {
                              "operator": ":",
                              "operands": [
                                {
                                  "operator": "==",
                                  "operands": [
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "@currentField"
                                      ]
                                    },
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "有給休暇"
                                      ]
                                    }
                                  ]
                                },
                                "",
                                {
                                  "operator": ":",
                                  "operands": [
                                    {
                                      "operator": "==",
                                      "operands": [
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "@currentField"
                                          ]
                                        },
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "有給休暇（午前）"
                                          ]
                                        }
                                      ]
                                    },
                                    "",
                                    {
                                      "operator": ":",
                                      "operands": [
                                        {
                                          "operator": "==",
                                          "operands": [
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "@currentField"
                                              ]
                                            },
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "有給休暇（午後）"
                                              ]
                                            }
                                          ]
                                        },
                                        "",
                                        {
                                          "operator": ":",
                                          "operands": [
                                            {
                                              "operator": "==",
                                              "operands": [
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    "@currentField"
                                                  ]
                                                },
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    ""
                                                  ]
                                                }
                                              ]
                                            },
                                            "",
                                            ""
                                          ]
                                        }
                                      ]
                                    }
                                  ]
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        "class": {
          "operator": ":",
          "operands": [
            {
              "operator": "==",
              "operands": [
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "@currentField"
                  ]
                },
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "在宅（勤務）"
                  ]
                }
              ]
            },
            "",
            {
              "operator": ":",
              "operands": [
                {
                  "operator": "==",
                  "operands": [
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "@currentField"
                      ]
                    },
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "在宅（待機）"
                      ]
                    }
                  ]
                },
                "",
                {
                  "operator": ":",
                  "operands": [
                    {
                      "operator": "==",
                      "operands": [
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "@currentField"
                          ]
                        },
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "出社"
                          ]
                        }
                      ]
                    },
                    "",
                    {
                      "operator": ":",
                      "operands": [
                        {
                          "operator": "==",
                          "operands": [
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "@currentField"
                              ]
                            },
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "外出"
                              ]
                            }
                          ]
                        },
                        "",
                        {
                          "operator": ":",
                          "operands": [
                            {
                              "operator": "==",
                              "operands": [
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "@currentField"
                                  ]
                                },
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "休み"
                                  ]
                                }
                              ]
                            },
                            "",
                            {
                              "operator": ":",
                              "operands": [
                                {
                                  "operator": "==",
                                  "operands": [
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "@currentField"
                                      ]
                                    },
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "有給休暇"
                                      ]
                                    }
                                  ]
                                },
                                "",
                                {
                                  "operator": ":",
                                  "operands": [
                                    {
                                      "operator": "==",
                                      "operands": [
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "@currentField"
                                          ]
                                        },
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "有給休暇（午前）"
                                          ]
                                        }
                                      ]
                                    },
                                    "",
                                    {
                                      "operator": ":",
                                      "operands": [
                                        {
                                          "operator": "==",
                                          "operands": [
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "@currentField"
                                              ]
                                            },
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "有給休暇（午後）"
                                              ]
                                            }
                                          ]
                                        },
                                        "",
                                        {
                                          "operator": ":",
                                          "operands": [
                                            {
                                              "operator": "==",
                                              "operands": [
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    "@currentField"
                                                  ]
                                                },
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    ""
                                                  ]
                                                }
                                              ]
                                            },
                                            "",
                                            ""
                                          ]
                                        }
                                      ]
                                    }
                                  ]
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      }
    },
    {
      "elmType": "span",
      "style": {
        "overflow": "hidden",
        "text-overflow": "ellipsis",
        "padding": "0 2px"
      },
      "txtContent": "@currentField",
      "attributes": {
        "class": {
          "operator": ":",
          "operands": [
            {
              "operator": "==",
              "operands": [
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "@currentField"
                  ]
                },
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "在宅（勤務）"
                  ]
                }
              ]
            },
            "",
            {
              "operator": ":",
              "operands": [
                {
                  "operator": "==",
                  "operands": [
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "@currentField"
                      ]
                    },
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "在宅（待機）"
                      ]
                    }
                  ]
                },
                "",
                {
                  "operator": ":",
                  "operands": [
                    {
                      "operator": "==",
                      "operands": [
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "@currentField"
                          ]
                        },
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "出社"
                          ]
                        }
                      ]
                    },
                    "",
                    {
                      "operator": ":",
                      "operands": [
                        {
                          "operator": "==",
                          "operands": [
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "@currentField"
                              ]
                            },
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "外出"
                              ]
                            }
                          ]
                        },
                        "",
                        {
                          "operator": ":",
                          "operands": [
                            {
                              "operator": "==",
                              "operands": [
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "@currentField"
                                  ]
                                },
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "休み"
                                  ]
                                }
                              ]
                            },
                            "",
                            {
                              "operator": ":",
                              "operands": [
                                {
                                  "operator": "==",
                                  "operands": [
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "@currentField"
                                      ]
                                    },
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "有給休暇"
                                      ]
                                    }
                                  ]
                                },
                                "",
                                {
                                  "operator": ":",
                                  "operands": [
                                    {
                                      "operator": "==",
                                      "operands": [
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "@currentField"
                                          ]
                                        },
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "有給休暇（午前）"
                                          ]
                                        }
                                      ]
                                    },
                                    "",
                                    {
                                      "operator": ":",
                                      "operands": [
                                        {
                                          "operator": "==",
                                          "operands": [
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "@currentField"
                                              ]
                                            },
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "有給休暇（午後）"
                                              ]
                                            }
                                          ]
                                        },
                                        "",
                                        {
                                          "operator": ":",
                                          "operands": [
                                            {
                                              "operator": "==",
                                              "operands": [
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    "@currentField"
                                                  ]
                                                },
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    ""
                                                  ]
                                                }
                                              ]
                                            },
                                            "",
                                            ""
                                          ]
                                        }
                                      ]
                                    }
                                  ]
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      }
    }
  ]
}
"@

$JsonFormatSunday = @"
{
  "`$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "style": {
    "box-sizing": "border-box",
    "padding": "0 2px"
  },
  "attributes": {
    "class": {
      "operator": ":",
      "operands": [
        {
          "operator": "==",
          "operands": [
            {
              "operator": "toLowerCase",
              "operands": [
                "@currentField"
              ]
            },
            {
              "operator": "toLowerCase",
              "operands": [
                "在宅（勤務）"
              ]
            }
          ]
        },
        "sp-css-backgroundColor-blueBackground37",
        {
          "operator": ":",
          "operands": [
            {
              "operator": "==",
              "operands": [
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "@currentField"
                  ]
                },
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "在宅（待機）"
                  ]
                }
              ]
            },
            "sp-css-backgroundColor-blueBackground27",
            {
              "operator": ":",
              "operands": [
                {
                  "operator": "==",
                  "operands": [
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "@currentField"
                      ]
                    },
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "出社"
                      ]
                    }
                  ]
                },
                "sp-css-backgroundColor-blueBackground17",
                {
                  "operator": ":",
                  "operands": [
                    {
                      "operator": "==",
                      "operands": [
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "@currentField"
                          ]
                        },
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "外出"
                          ]
                        }
                      ]
                    },
                    "sp-css-backgroundColor-blueBackground07",
                    {
                      "operator": ":",
                      "operands": [
                        {
                          "operator": "==",
                          "operands": [
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "@currentField"
                              ]
                            },
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "休み"
                              ]
                            }
                          ]
                        },
                        "sp-css-backgroundColor-successBackground50",
                        {
                          "operator": ":",
                          "operands": [
                            {
                              "operator": "==",
                              "operands": [
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "@currentField"
                                  ]
                                },
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "有給休暇"
                                  ]
                                }
                              ]
                            },
                            "sp-css-backgroundColor-successBackground40",
                            {
                              "operator": ":",
                              "operands": [
                                {
                                  "operator": "==",
                                  "operands": [
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "@currentField"
                                      ]
                                    },
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "有給休暇（午前）"
                                      ]
                                    }
                                  ]
                                },
                                "sp-css-backgroundColor-successBackground30",
                                {
                                  "operator": ":",
                                  "operands": [
                                    {
                                      "operator": "==",
                                      "operands": [
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "@currentField"
                                          ]
                                        },
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "有給休暇（午後）"
                                          ]
                                        }
                                      ]
                                    },
                                    "sp-css-backgroundColor-successBackground",
                                    {
                                      "operator": ":",
                                      "operands": [
                                        {
                                          "operator": "==",
                                          "operands": [
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "@currentField"
                                              ]
                                            },
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                ""
                                              ]
                                            }
                                          ]
                                        },
                                        "sp-css-backgroundColor-errorBackground",
                                        ""
                                      ]
                                    }
                                  ]
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    }
  },
  "children": [
    {
      "elmType": "span",
      "attributes": {
        "iconName": {
          "operator": ":",
          "operands": [
            {
              "operator": "==",
              "operands": [
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "@currentField"
                  ]
                },
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "在宅（勤務）"
                  ]
                }
              ]
            },
            "",
            {
              "operator": ":",
              "operands": [
                {
                  "operator": "==",
                  "operands": [
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "@currentField"
                      ]
                    },
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "在宅（待機）"
                      ]
                    }
                  ]
                },
                "",
                {
                  "operator": ":",
                  "operands": [
                    {
                      "operator": "==",
                      "operands": [
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "@currentField"
                          ]
                        },
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "出社"
                          ]
                        }
                      ]
                    },
                    "",
                    {
                      "operator": ":",
                      "operands": [
                        {
                          "operator": "==",
                          "operands": [
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "@currentField"
                              ]
                            },
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "外出"
                              ]
                            }
                          ]
                        },
                        "",
                        {
                          "operator": ":",
                          "operands": [
                            {
                              "operator": "==",
                              "operands": [
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "@currentField"
                                  ]
                                },
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "休み"
                                  ]
                                }
                              ]
                            },
                            "",
                            {
                              "operator": ":",
                              "operands": [
                                {
                                  "operator": "==",
                                  "operands": [
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "@currentField"
                                      ]
                                    },
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "有給休暇"
                                      ]
                                    }
                                  ]
                                },
                                "",
                                {
                                  "operator": ":",
                                  "operands": [
                                    {
                                      "operator": "==",
                                      "operands": [
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "@currentField"
                                          ]
                                        },
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "有給休暇（午前）"
                                          ]
                                        }
                                      ]
                                    },
                                    "",
                                    {
                                      "operator": ":",
                                      "operands": [
                                        {
                                          "operator": "==",
                                          "operands": [
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "@currentField"
                                              ]
                                            },
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "有給休暇（午後）"
                                              ]
                                            }
                                          ]
                                        },
                                        "",
                                        {
                                          "operator": ":",
                                          "operands": [
                                            {
                                              "operator": "==",
                                              "operands": [
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    "@currentField"
                                                  ]
                                                },
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    ""
                                                  ]
                                                }
                                              ]
                                            },
                                            "",
                                            ""
                                          ]
                                        }
                                      ]
                                    }
                                  ]
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        "class": {
          "operator": ":",
          "operands": [
            {
              "operator": "==",
              "operands": [
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "@currentField"
                  ]
                },
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "在宅（勤務）"
                  ]
                }
              ]
            },
            "",
            {
              "operator": ":",
              "operands": [
                {
                  "operator": "==",
                  "operands": [
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "@currentField"
                      ]
                    },
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "在宅（待機）"
                      ]
                    }
                  ]
                },
                "",
                {
                  "operator": ":",
                  "operands": [
                    {
                      "operator": "==",
                      "operands": [
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "@currentField"
                          ]
                        },
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "出社"
                          ]
                        }
                      ]
                    },
                    "",
                    {
                      "operator": ":",
                      "operands": [
                        {
                          "operator": "==",
                          "operands": [
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "@currentField"
                              ]
                            },
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "外出"
                              ]
                            }
                          ]
                        },
                        "",
                        {
                          "operator": ":",
                          "operands": [
                            {
                              "operator": "==",
                              "operands": [
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "@currentField"
                                  ]
                                },
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "休み"
                                  ]
                                }
                              ]
                            },
                            "",
                            {
                              "operator": ":",
                              "operands": [
                                {
                                  "operator": "==",
                                  "operands": [
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "@currentField"
                                      ]
                                    },
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "有給休暇"
                                      ]
                                    }
                                  ]
                                },
                                "",
                                {
                                  "operator": ":",
                                  "operands": [
                                    {
                                      "operator": "==",
                                      "operands": [
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "@currentField"
                                          ]
                                        },
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "有給休暇（午前）"
                                          ]
                                        }
                                      ]
                                    },
                                    "",
                                    {
                                      "operator": ":",
                                      "operands": [
                                        {
                                          "operator": "==",
                                          "operands": [
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "@currentField"
                                              ]
                                            },
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "有給休暇（午後）"
                                              ]
                                            }
                                          ]
                                        },
                                        "",
                                        {
                                          "operator": ":",
                                          "operands": [
                                            {
                                              "operator": "==",
                                              "operands": [
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    "@currentField"
                                                  ]
                                                },
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    ""
                                                  ]
                                                }
                                              ]
                                            },
                                            "",
                                            ""
                                          ]
                                        }
                                      ]
                                    }
                                  ]
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      }
    },
    {
      "elmType": "span",
      "style": {
        "overflow": "hidden",
        "text-overflow": "ellipsis",
        "padding": "0 2px"
      },
      "txtContent": "@currentField",
      "attributes": {
        "class": {
          "operator": ":",
          "operands": [
            {
              "operator": "==",
              "operands": [
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "@currentField"
                  ]
                },
                {
                  "operator": "toLowerCase",
                  "operands": [
                    "在宅（勤務）"
                  ]
                }
              ]
            },
            "",
            {
              "operator": ":",
              "operands": [
                {
                  "operator": "==",
                  "operands": [
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "@currentField"
                      ]
                    },
                    {
                      "operator": "toLowerCase",
                      "operands": [
                        "在宅（待機）"
                      ]
                    }
                  ]
                },
                "",
                {
                  "operator": ":",
                  "operands": [
                    {
                      "operator": "==",
                      "operands": [
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "@currentField"
                          ]
                        },
                        {
                          "operator": "toLowerCase",
                          "operands": [
                            "出社"
                          ]
                        }
                      ]
                    },
                    "",
                    {
                      "operator": ":",
                      "operands": [
                        {
                          "operator": "==",
                          "operands": [
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "@currentField"
                              ]
                            },
                            {
                              "operator": "toLowerCase",
                              "operands": [
                                "外出"
                              ]
                            }
                          ]
                        },
                        "",
                        {
                          "operator": ":",
                          "operands": [
                            {
                              "operator": "==",
                              "operands": [
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "@currentField"
                                  ]
                                },
                                {
                                  "operator": "toLowerCase",
                                  "operands": [
                                    "休み"
                                  ]
                                }
                              ]
                            },
                            "",
                            {
                              "operator": ":",
                              "operands": [
                                {
                                  "operator": "==",
                                  "operands": [
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "@currentField"
                                      ]
                                    },
                                    {
                                      "operator": "toLowerCase",
                                      "operands": [
                                        "有給休暇"
                                      ]
                                    }
                                  ]
                                },
                                "",
                                {
                                  "operator": ":",
                                  "operands": [
                                    {
                                      "operator": "==",
                                      "operands": [
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "@currentField"
                                          ]
                                        },
                                        {
                                          "operator": "toLowerCase",
                                          "operands": [
                                            "有給休暇（午前）"
                                          ]
                                        }
                                      ]
                                    },
                                    "",
                                    {
                                      "operator": ":",
                                      "operands": [
                                        {
                                          "operator": "==",
                                          "operands": [
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "@currentField"
                                              ]
                                            },
                                            {
                                              "operator": "toLowerCase",
                                              "operands": [
                                                "有給休暇（午後）"
                                              ]
                                            }
                                          ]
                                        },
                                        "",
                                        {
                                          "operator": ":",
                                          "operands": [
                                            {
                                              "operator": "==",
                                              "operands": [
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    "@currentField"
                                                  ]
                                                },
                                                {
                                                  "operator": "toLowerCase",
                                                  "operands": [
                                                    ""
                                                  ]
                                                }
                                              ]
                                            },
                                            "",
                                            ""
                                          ]
                                        }
                                      ]
                                    }
                                  ]
                                }
                              ]
                            }
                          ]
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      }
    }
  ]
}
"@

$Fields = $List.Fields
$Context.Load($Fields)
$Context.ExecuteQuery()

$SatuadayFields = $List.Fields | Where-Object { $_.Title.Contains("(土)") }
Foreach ($Field in $SatuadayFields) {
    $Context.Load($Field)
    $Context.ExecuteQuery()
 
    $Field.CustomFormatter = $JsonFormatSaturday
    $Field.Update()
    $Context.ExecuteQuery()   
}

$SundayFields = $List.Fields | Where-Object { $_.Title.Contains("(日)") }
Foreach ($Field in $SundayFields) {
    $Context.Load($Field)
    $Context.ExecuteQuery()
 
    $Field.CustomFormatter = $JsonFormatSunday
    $Field.Update()
    $Context.ExecuteQuery()   
}

$Views = $List.Views
$Context.Load($Views)
$Context.ExecuteQuery()
Log("$List.Views Load Executed")

$TargetString = $TargetDate.AddMonths(-1).ToString("_yyyy年MM月") 

$TargetViews = $Views | Where-Object {$_.Title.Contains($TargetString)}

ForEach($TargetView in $TargetViews)
{
    if($Null -eq $TargetView -or $Null -eq $TargetView.Title) {
      continue
    }

    $TargetDate = $InitialDate
    $SplitStrings = $TargetView.Title.Split("_")
    $Department = $SplitStrings[0]

    $TargetYearMonth = $TargetDate.ToString("yyyy") + "年" + $TargetDate.ToString("MM") + "月"
    $ViewCreationInfo = New-Object Microsoft.SharePoint.Client.ViewCreationInformation
    $ViewCreationInfo.Title = $Department + "_" + $TargetYearMonth
    $ViewCreationInfo.Query = $TargetView.ViewQuery
    $ViewCreationInfo.RowLimit = "300"

    $FieldArray = @("名前","部署")
    For ($TargetDate; $TargetDate -le $LastDate; $TargetDate = $TargetDate.AddDays(1)) {
        $Day = $TargetDate.ToString("M/d (ddd)")
        $FieldArray += $Day
    }

    $ViewCreationInfo.ViewFields = $FieldArray
    if ($ViewCreationInfo.Title.Equals("All_" + $TargetYearMonth)) {
        $ViewCreationInfo.SetAsDefaultView = $True
    }
    else {
        $ViewCreationInfo.SetAsDefaultView = $False
    }

    $NewView = $Views.Add($ViewCreationInfo)
    $Context.ExecuteQuery()
}

$Context.Dispose()
Disconnect-SPOService
Log("Disconnect-SPOService")