//
// Banner.swift
// Bridge Score
//
//  Created by Marc Shearer on 05/02/2021.
//

import SwiftUI

struct BannerOption {
    let image: AnyView?
    let text: String?
    let likeBack: Bool
    let color: PaletteColor?
    @Binding var isEnabled: Bool
    @Binding var isHidden: Bool
    let menu: Bool
    let action: ()->()
    
    init(image: AnyView? = nil, text: String? = nil, color: PaletteColor? = nil, likeBack: Bool = false, isEnabled: Binding<Bool>? = nil, isHidden: Binding<Bool>? = nil, menu: Bool = false, action: @escaping ()->()) {
        self.image = image
        self.text = text
        self.likeBack = likeBack
        self.color = color
        self.action = action
        self._isEnabled = isEnabled ?? Binding.constant(true)
        self._isHidden = isHidden ?? Binding.constant(false)
        self.menu = menu
    }
}

enum BannerOptionMode {
    case menu
    case buttons
    case both
    case none
}

struct Banner: View {
    @Environment(\.presentationMode) private var presentationMode: Binding<PresentationMode>

    @Binding var title: String
    var alternateStyle: Bool = false
    var bottomSpace: Bool = true
    var back: Bool = true
    var backEnabled: (()->(Bool))?
    var backText: String? = nil
    var backImage: AnyView? = AnyView(Image(systemName: "chevron.left"))
    var backAction: (()->(Bool))?
    var leftTitle: Bool?
    var optionMode: BannerOptionMode = .none
    var menuImage: AnyView? = nil
    var menuTitle: String?
    var menuId: UUID? = nil
    var options: [BannerOption]? = nil
    var disabled: Binding<Bool> = Binding.constant(false)
    public static let crossImage = AnyView(Image(systemName: "xmark"))
    
    @State private var bannerColor: PaletteColor = Palette.banner
    @State private var buttonColor: PaletteColor = Palette.bannerButton
    @State private var backButtonColor: Color = Palette.bannerBackButton
    
    var body: some View {
        return ZStack {
            bannerColor.background
                .ignoresSafeArea(edges: .all)
            VStack {
                Spacer()
                HStack {
                    Spacer().frame(width: 20)

                    if (leftTitle ?? !back) {
                        HStack {
                            backButton
                            if MyApp.format != .phone || isLandscape || (optionMode != .buttons && optionMode != .both) {
                                Spacer().frame(width: 12)
                                titleText
                                .onTapGesture {
                                    backPressed()
                                }
                                Spacer().frame(width: 20)
                            }
                            Spacer()
                            menu
                        }
                        
                    } else {
                        ZStack {
                            HStack {
                                backButton
                                Spacer()
                                menu
                            }
                            HStack {
                                Spacer().frame(width: 80)
                                Spacer()
                                if MyApp.format != .phone || isLandscape || (optionMode != .buttons && optionMode != .both) {
                                    titleText
                                    Spacer()
                                }
                                Spacer().frame(width: 80)
                            }
                        }
                    }
                
                    Spacer().frame(width: 20)
                }
                Spacer().frame(height: (alternateStyle ? 0 : bannerBottom))
            }
        }
        .disabled(disabled.wrappedValue)
        .onAppear {
            bannerColor = (alternateStyle ? Palette.alternateBanner : Palette.banner)
            buttonColor = (alternateStyle ? Palette.alternateBannerButton : Palette.bannerButton)
            backButtonColor = (alternateStyle ? Palette.alternateBannerBackButton : Palette.bannerBackButton)
        }
        .frame(height: (alternateStyle ? alternateBannerHeight : bannerHeight + bannerBottom))
        .background(bannerColor.background)
    }
        
    var backButton: some View {
        HStack {
            if back {
                let enabled = backEnabled?() ?? true
                Button(action: {
                    backPressed()
                }, label: {
                    HStack {
                        if let backText = backText {
                            Text(backText)
                                .font((alternateStyle ? alternateBannerFont : bannerFont)).bold()
                        } else {
                            backImage
                                .font(bannerFont)
                        }
                        
                    }
                    .foregroundColor(backButtonColor.opacity(enabled ? 1.0 : 0.5))
                })
                .disabled(!(enabled))
            } else {
                EmptyView()
            }
        }
    }
    
    func backPressed() {
        if backEnabled?() ?? true {
            if backAction?() ?? true {
                self.presentationMode.wrappedValue.dismiss()
            }
        }
    }
    
    var titleText: some View {
        Text(title)
            .font((alternateStyle ? alternateBannerFont : bannerFont)).bold()
            .foregroundColor(bannerColor.text)
    }
    
    var menu: some View {
        HStack {
            if let options = options {
                if optionMode == .buttons || optionMode == .both {
                    Banner_Buttons(options: options.filter{!$0.menu || optionMode == .buttons}, alternateStyle: alternateStyle, bannerColor: bannerColor, buttonColor: buttonColor, backButtonColor: backButtonColor)
                }
                if optionMode == .menu || optionMode == .both {
                    Banner_Menu(id: menuId!, image: menuImage, title: menuTitle, options: options.filter{$0.menu || optionMode == .menu}, bannerColor: bannerColor)
                }
            }
        }
    }
}

struct Banner_Menu : View {
    var id: UUID
    var image: AnyView?
    var title: String?
    var options: [BannerOption]
    var bannerColor: PaletteColor
    let menuStyle = DefaultMenuStyle()

    var body: some View {
        Button {
            let filteredOptions = options.filter{$0.isEnabled && !$0.isHidden}
            SlideInMenu.shared.show(id: id, title: title, strings: filteredOptions.map{$0.text ?? ""}, top: bannerHeight - 20) { (option) in
                    if let selected = options.first(where: {$0.text == option}) {
                        selected.action()
                    }
                }
            } label: {
                (image ?? AnyView(Image(systemName: "line.3.horizontal"))).foregroundColor(bannerColor.text).font(bannerFont)
            }
    
    }
}

struct Banner_Buttons : View {
    var options: [BannerOption]
    var alternateStyle: Bool
    var bannerColor: PaletteColor
    var buttonColor: PaletteColor
    var backButtonColor: Color
    
    var body: some View {
        HStack {
            ForEach(options.indices, id: \.self) { (index) in
                let option = options[index]
                let backgroundColor = option.color?.background ?? (option.likeBack ? bannerColor.background : buttonColor.background)
                let foregroundColor = (option.isEnabled ?
                                        (option.color?.text ?? (option.likeBack ? backButtonColor : buttonColor.text)) :
                                        (option.color ?? buttonColor).faintText)
                if !option.isHidden {
                    HStack {
                        Button {
                            option.action()
                        } label: {
                            VStack {
                                if !option.likeBack {
                                    Spacer().frame(height: 6)
                                }
                                HStack {
                                    if !option.likeBack {
                                        Spacer().frame(width: 16)
                                    }
                                    if option.image != nil {
                                        option.image.foregroundColor(foregroundColor)
                                    }
                                    if option.image != nil && option.text != nil {
                                        Spacer().frame(width: 16)
                                    }
                                    if option.text != nil {
                                        Text(option.text ?? "").foregroundColor(foregroundColor)
                                    }
                                    if !option.likeBack {
                                        Spacer().frame(width: 16)
                                    }
                                }
                                if !option.likeBack {
                                    Spacer().frame(height: 6)
                                }
                            }
                        }
                        .disabled(!option.isEnabled)
                        .font(alternateStyle || MyApp.format == .phone ? alternateBannerFont : bannerFont)
                        .background(backgroundColor)
                        .cornerRadius(option.likeBack ? 0 : 10.0)
                        if index != options.count - 1 {
                            Spacer().frame(width: 16)
                        }
                    }
                }
            }
        }
    }
}
