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
                                    
                                    InputFloat(title: "Points award:", field: $settings.maxPoints, topSpace: 0, leadingSpace: 30, width: 50, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    Spacer()
                                }
                            }
                                
                            VStack {
                                InputTitle(title: "Good Status Values:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    
                                   Input(title: "Status  :", field: $settings.goodStatus, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                    
                                    Input(title: "Payment  :", field: $settings.goodPaymentStatus, topSpace: 0, leadingSpace: 30, width: 250, inlineTitle: true, inlineTitleWidth: 120, isEnabled: true)
                                    
                                    Spacer()
                                }
                                
                                Spacer().frame(height: 10)
                                
                                HStack {
                                    Spacer().frame(width: 220)
                                    
                                    InputInt(title: "Ignore months :", field: $settings.ignorePaymentFrom, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 114, isEnabled: true)
                                    
                                    InputInt(title: "to :", field: $settings.ignorePaymentTo, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 30, isEnabled: true)
                                    
                                    Spacer()
                                }
                            }
            
                            VStack {
                                InputTitle(title: "Maximum sizes:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    
                                    InputInt(title: "Field entry:", field: $settings.largestFieldSize, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    InputInt(title: "Players:", field: $settings.largestPlayerCount, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                InputTitle(title: "Presentation:", topSpace: 16)
                                Spacer().frame(height: 8)
                                HStack {
                                    
                                    InputFloat(title: "Default zoom:", field: $settings.defaultWorksheetZoom, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    InputInt(title: "Lines/page:", field: $settings.linesPerFormattedPage, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 100)
                                    
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                InputTitle(title: "User Download Details:", topSpace: 16)
                                
                                HStack {
                                    Input(title: "Filename:", field: $settings.userDownloadData, leadingSpace: 30, width: 160, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                    
                                    InputInt(title: "First Row:", field: $settings.userDownloadMinRow, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 80)
                                    
                                    InputInt(title: "Last Row:", field: $settings.userDownloadMaxRow, topSpace: 0, leadingSpace: 30, width: 80, inlineTitle: true, inlineTitleWidth: 80)
                                    
                                    Spacer()
                                }
                                HStack {
                                    Input(title: "Frozen file:", field: $settings.userDownloadFrozenData, leadingSpace: 30, width: 160, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                    Spacer()
                                }
                            }
                            
                            VStack {
                                Spacer().frame(height: 8)
                                
                                VStack {
                                    InputTitle(title: "User Download Columns:", topSpace: 16)
                                    
                                    HStack {
                                        
                                        Input(title: "National ID:", field: $settings.userDownloadNationalIdColumn, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "First Name:", field: $settings.userDownloadFirstNameColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Other Names:", field: $settings.userDownloadOtherNamesColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Rank:", field: $settings.userDownloadRankColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Spacer()
                                    }
                                    
                                    HStack {
                                        
                                        Input(title: "Email:", field: $settings.userDownloadEmailColumn, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Home Club:", field: $settings.userDownloadHomeClubColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Status:", field: $settings.userDownloadStatusColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Input(title: "Payment:", field: $settings.userDownloadPaymentStatusColumn, topSpace: 0, leadingSpace: 12, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
                                        Spacer()
                                        
                                    }
                                    
                                    HStack {
                                        
                                        Input(title: "Other Union:", field: $settings.userDownloadOtherUnionColumn, topSpace: 0, leadingSpace: 30, width: 40, inlineTitle: true, inlineTitleWidth: 100, isEnabled: true)
                                        
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
