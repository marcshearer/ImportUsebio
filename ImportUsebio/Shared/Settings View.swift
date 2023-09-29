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
                })
                HStack {
                    Spacer().frame(width: 30)
                    VStack {
                        VStack {
                            
                            VStack {
                                InputTitle(title: "Maximum value warnings:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    Spacer().frame(width: 30)
                                    
                                    InputInt(title: "National Id number:", field:$settings.maxNationalIdNumber, topSpace: 0, width: 80, inlineTitle: true, inlineTitleWidth: 140)
                                    
                                    InputFloat(title: "Points award:", field: $settings.maxPoints, topSpace: 0, width: 50, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    Spacer()
                                }
                            }
                                
                            HStack {
                                Input(title: "Good member status:", field: $settings.goodStatus, width: 300, isEnabled: true)
                                Spacer()
                            }
            
                            VStack {
                                InputTitle(title: "Maximum sizes:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    Spacer().frame(width: 30)
                                    
                                    InputInt(title: "Field entry:", field: $settings.largestFieldSize, topSpace: 0, width: 80, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    InputInt(title: "Players:", field: $settings.largestPlayerCount, topSpace: 0, width: 80, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                InputTitle(title: "Presentation:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    Spacer().frame(width: 30)
                                    
                                    InputFloat(title: "Default zoom:", field: $settings.defaultWorksheetZoom, topSpace: 0, width: 80, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    InputInt(title: "Lines per formatted page:", field: $settings.linesPerFormattedPage, topSpace: 0, width: 80, inlineTitle: true, inlineTitleWidth: 180)
                                    
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                InputTitle(title: "User Download Details:", topSpace: 16)
                                
                                HStack {
                                    Input(title: "Filename:", field: $settings.userDownloadData, leadingSpace: 42, width: 180, inlineTitle: true, inlineTitleWidth: 80, isEnabled: true)
                                    
                                    InputInt(title: "First Row:", field: $settings.userDownloadMinRow, topSpace: 0, width: 80, inlineTitle: true, inlineTitleWidth: 80)
                                    
                                    InputInt(title: "Last Row:", field: $settings.userDownloadMaxRow, topSpace: 0, width: 80, inlineTitle: true, inlineTitleWidth: 80)
                                    
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                Spacer().frame(height: 8)
                                
                                VStack {
                                    InputTitle(title: "User Download Columns:", topSpace: 16)
                                    
                                    HStack {
                                        
                                        Input(title: "National ID:", field: $settings.userDownloadNationalIdColumn, topSpace: 0, leadingSpace: 42, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "First Name:", field: $settings.userDownloadFirstNameColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Other Names:", field: $settings.userDownloadOtherNamesColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Rank:", field: $settings.userDownloadRankColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Spacer()
                                    }
                                    
                                    HStack {
                                        
                                        Input(title: "Email:", field: $settings.userDownloadEmailColumn, topSpace: 0, leadingSpace: 42, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Home Club:", field: $settings.userDownloadHomeClubColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Status:", field: $settings.userDownloadStatusColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Other Union:", field: $settings.userDownloadOtherUnionColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
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
        .frame(width: 800, height: 580)
    }
}
