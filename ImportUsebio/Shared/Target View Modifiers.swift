//
//  Target View Modifiers.swift
// Bridge Score
//
//  Created by Marc Shearer on 26/02/2021.
//

import SwiftUI

struct NoNavigationBar : ViewModifier {
        
    #if canImport(UIKit)
    func body(content: Content) -> some View { content
        .navigationBarBackButtonHidden(true)
        .navigationBarTitle("")
        .navigationBarHidden(true)
    }
    #else
    func body(content: Content) -> some View { content
        
    }
    #endif
}

extension View {
    var noNavigationBar: some View {
        self.modifier(NoNavigationBar())
    }
}

struct RightSpacer : ViewModifier {
        
    func body(content: Content) -> some View { content
        .frame(width: (MyApp.target == .iOS ? 16 : 32))
    }
}

extension View {
    var rightSpacer: some View {
        self.modifier(RightSpacer())
    }
}

struct BottomSpacer : ViewModifier {
        
    func body(content: Content) -> some View { content
        .frame(height: (MyApp.target == .iOS ? 0 : 16))
    }
}

extension View {
    var bottomSpacer: some View {
        self.modifier(BottomSpacer())
    }
}

struct MyEditModeModifier : ViewModifier {
    @Binding var editMode: MyEditMode
    #if canImport(UIKit)
    func body(content: Content) -> some View { content
        .environment(\.editMode, $editMode)
    }
    #else
    func body(content: Content) -> some View { content
        
    }
    #endif
}

extension View {
    func editMode(_ editMode: Binding<MyEditMode>) -> some View {
        self.modifier(MyEditModeModifier(editMode: editMode))
    }
}

#if canImport(UIKit)
typealias MyEditMode = EditMode
#else
typealias MyEditMode = Bool
#endif

#if canImport(UIKit)
typealias IosStackNavigationViewStyle = StackNavigationViewStyle
#else
typealias IosStackNavigationViewStyle = DefaultNavigationViewStyle
#endif

struct MySheetViewModifier<Item, SheetContent : View> : ViewModifier where Item : Identifiable {
    let item: Binding<Item?>
    let onDismiss: (()->())?
    let sheetContent: (Item)->(SheetContent)
    
    init(item: Binding<Item?>, onDismiss: (()->())? = nil, @ViewBuilder content: @escaping (Item)->(SheetContent)) {
        self.item = item
        self.onDismiss = onDismiss
        self.sheetContent = content
    }
    
    #if canImport(UIKit)
    func body(content: Content) -> some View {
        content.fullScreenCover(item: item, onDismiss: onDismiss) { (item) in
            sheetContent(item)
        }
    }
    #else
    func body(content: Content) -> some View {
        content.sheet(item: item, onDismiss: onDismiss) { (item) in
            sheetContent(item)
        }
    }
    #endif
}

extension View {
    func mySheet<Item, SheetContent>(item: Binding<Item?>, onDismiss: (()->())? = nil, @ViewBuilder content: @escaping (Item)->(SheetContent)) -> some View where Item : Identifiable, SheetContent: View {
        self.modifier(MySheetViewModifier(item: item, onDismiss: onDismiss, content: content))
    }
}

struct MyKeyboardTypeViewModifier : ViewModifier {
    @State var keyboardType: KeyboardType
    #if canImport(UIKit)
    func body(content: Content) -> some View { content
        .keyboardType(keyboardType)
    }
    #else
    func body(content: Content) -> some View { content
        
    }
    #endif
}

extension View {
    func myKeyboardType(_ keyboardType: KeyboardType) -> some View {
        self.modifier(MyKeyboardTypeViewModifier(keyboardType: keyboardType))
    }
}

#if canImport(UIKit)
typealias KeyboardType = UIKeyboardType
#else
enum KeyboardType {
    case `default`
    case URL
}
#endif

struct MyAutoCapitalizationViewModifier : ViewModifier {
    @State var autocapitalization: AutoCapitalization
    #if canImport(UIKit)
    func body(content: Content) -> some View { content
        .autocapitalization(autocapitalization)
    }
    #else
    func body(content: Content) -> some View { content
        
    }
    #endif
}

extension View {
    func myAutocapitalization(_ autoCapitalization: AutoCapitalization) -> some View {
        self.modifier(MyAutoCapitalizationViewModifier(autocapitalization: autoCapitalization))
    }
}

#if canImport(UIKit)
typealias AutoCapitalization = UITextAutocapitalizationType
#else
enum AutoCapitalization {
    case sentences
    case allCharacters
    case none
}
#endif
