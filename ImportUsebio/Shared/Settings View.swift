//
//  Settings.swift
//  ImportUsebio
//
//  Created by Marc Shearer on 20/03/2023.
//

import SwiftUI

struct SettingsView: View {
    @ObservedObject var settings: Settings
    
    @State private var refresh = true
    
    var body: some View {
        // Just to trigger view refresh
        if refresh { EmptyView() }
            
        StandardView("Select Input") {
            VStack {
                Banner(title: Binding.constant("Settings"), backAction: {
                    Settings.current.copy(from: settings)
                    Settings.current.save()
                    return true
                }, optionMode: .buttons, options: [versionOption])
                HStack {
                    Spacer().frame(width: 30)
                    VStack {
                        VStack {
                            
                            VStack {
                                InputTitle(title: "Maximum value warnings:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    
                                    InputInt(title: "SBU Number:", field:$settings.maxNationalIdNumber, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 100)
                                        .help("SBU Numbers above this value will be highlighted")
                                    
                                    InputFloat(title: "Points award:", field: $settings.maxPoints, topSpace: 0, leadingSpace: 30, width: 50, inlineTitle: true, inlineTitleWidth: 100)
                                        .help("Point awards to a player above this will be highlighted")
                                    
                                    Spacer()
                                }
                            }
                                
                            VStack {
                                InputTitle(title: "Specific Values:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    
                                   Input(title: "Status  :", field: $settings.goodStatus, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 110, isEnabled: true)
                                        .help("Value of status column in download that is considered 'Active'")
                                    
                                    Input(title: "Not Paid  :", field: $settings.notPaidPaymentStatus, topSpace: 0, leadingSpace: 30, width: 250, inlineTitle: true, inlineTitleWidth: 120, isEnabled: true)
                                        .help("Value in current season payment status column that is considered 'Not Paid'")
                                    
                                    Spacer()
                                }
                                
                                Spacer().frame(height: 10)
                                
                                HStack {
                                    
                                    InputInt(title: "Non-SBU rank: ", field: $settings.otherNBORank, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 108, isEnabled: true)
                                    
                                    Spacer().frame(width: 44)
                                    
                                    InputInt(title: "Ignore months :", field: $settings.ignorePaymentFrom, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 114, isEnabled: true)
                                        .help("Months after this month and before the 'to' month will be ignored for payment warning")
                                    
                                    InputInt(title: "to :", field: $settings.ignorePaymentTo, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 30, isEnabled: true)
                                        .help("Months up to this month after the 'from' month will be ignored for payment warning")
                                    
                                    Spacer()
                                }
                            }
            
                            VStack {
                                InputTitle(title: "Maximum sizes:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    
                                    InputInt(title: "Max round entry:", field: $settings.largestFieldSize, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 150)
                                        .help("Maximum number of participants (pairs/teams) in a round")
                                    
                                    InputInt(title: "Max round players:", field: $settings.largestPlayerCount, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 150)
                                        .help("Maximum number of players in a round")
                                    
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                InputTitle(title: "Presentation:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    
                                    InputFloat(title: "Default zoom:", field: $settings.defaultWorksheetZoom, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 150)
                                        .help("Default zoom factor of spreadsheet")
                                    
                                    InputInt(title: "Lines/page:", field: $settings.linesPerFormattedPage, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 150)
                                        .help("Default lines/page on the formatted tab of the spreadsheet")
                                    
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                InputTitle(title: "User Download Details:", topSpace: 16)
                                
                                HStack {
                                    Input(title: "Filename:", field: $settings.userDownloadData, leadingSpace: 30, width: 160, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        .help("Filename of the downloaded user data (including extension)")
                                    
                                    InputInt(title: "First Row:", field: $settings.userDownloadMinRow, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 80)
                                        .help("First data row in the downloaded user data")
                                    InputInt(title: "Max Row:", field: $settings.userDownloadMaxRow, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 80)
                                        .help("Maximum data row in the downloaded user data")
                                    
                                    Spacer()
                                }
                                HStack {
                                    Input(title: "Frozen file:", field: $settings.userDownloadFrozenData, leadingSpace: 30, width: 160, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        .help("Filename of the downloaded frozen user data (including extension)")
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                Spacer().frame(height: 8)
                                
                                VStack {
                                    InputTitle(title: "User Download Columns:", topSpace: 16)
                                    
                                    HStack {
                                        
                                        Input(title: "National ID:", field: $settings.userDownloadNationalIdColumn, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the national ID column in the downloaded data")
                                        
                                        Input(title: "First Name:", field: $settings.userDownloadFirstNameColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the first name column in the downloaded data")
                                        
                                        Input(title: "Other Names:", field: $settings.userDownloadOtherNamesColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the other names column in the downloaded data")
                                        
                                        Input(title: "Rank:", field: $settings.userDownloadRankColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the rank code column in the downloaded data")
                                        
                                        Spacer()
                                    }
                                    
                                    HStack {
                                        
                                        Input(title: "Email:", field: $settings.userDownloadEmailColumn, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the email column in the downloaded data")
                                        
                                        Input(title: "Home Club:", field: $settings.userDownloadHomeClubColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the home club column in the downloaded data")
                                        
                                        Input(title: "Status:", field: $settings.userDownloadStatusColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the status column in the downloaded data")
                                        
                                        Input(title: "Payment:", field: $settings.userDownloadPaymentStatusColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the current season payment status column in the downloaded data")
                                        
                                        Spacer()
                                        
                                    }
                                    
                                    HStack {
                                        
                                        Input(title: "Union (NBO):", field: $settings.userDownloadOtherUnionColumn, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                            .help("Column letter(s) of the union (NBO) column in the downloaded data")
                                        
                                        Spacer()
                                    }
                                }
                            }
                            
                            Spacer()
                        }
                    }
                    Spacer()
                }
            }
        }
        .frame(width: 800, height: 600)
    }
    
    private var versionOption: BannerOption {
        BannerOption(text: "Version: \(Version.current.version) (Build \(Version.current.build))", color: Palette.alternateBanner, likeBack: true, isEnabled: Binding.constant(false), action: {})
    }
    
}
