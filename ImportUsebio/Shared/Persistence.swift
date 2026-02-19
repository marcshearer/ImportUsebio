//
//  Persistence.swift
//  SystemControl
//
//  Created by Marc Shearer on 20/01/2022.
//

import CoreData

struct PersistenceController {
    static let shared = PersistenceController()

    static var preview: PersistenceController = {
        let result = PersistenceController(inMemory: true)
        let viewContext = result.container.viewContext
        
        // Setup preview data here
        
        do {
            try viewContext.save()
        } catch {
            // Replace this implementation with code to handle the error appropriately.
            // fatalError() causes the application to generate a crash log and terminate. You should not use this function in a shipping application, although it may be useful during development.
            let nsError = error as NSError
            fatalError("Unresolved error \(nsError), \(nsError.userInfo)")
        }
        return result
    }()

    let container: NSPersistentContainer

    init(inMemory: Bool = false) {
        container = NSPersistentContainer(name: "ImportUsebio", managedObjectModel: MyApp.objectModel.model)
          
        if inMemory {
            container.persistentStoreDescriptions.first!.url = URL(fileURLWithPath: "/dev/null")
        } else {
            /* Remove comment to synchronise with iCloud
               Also would need to re-create App Groups in entitlements
            container.viewContext.automaticallyMergesChangesFromParent = true
            // Get core data directory and append Development or Production
            let storeDirectory = FileManager.default.containerURL(forSecurityApplicationGroupIdentifier: appGroup)!
            
            // Create a store description for a local store
            let storeLocation = storeDirectory.appendingPathComponent("SystemControl-\(MyApp.expectedDatabase.name).sqlite")
            let storeDescription = NSPersistentStoreDescription(url: storeLocation)
            storeDescription.cloudKitContainerOptions =
                NSPersistentCloudKitContainerOptions(
                    containerIdentifier: iCloudIdentifier)
            
            let remoteChangeKey = "NSPersistentStoreRemoteChangeNotificationOptionKey"
            storeDescription.setOption(true as NSNumber, forKey: remoteChangeKey)
            let historyTrackingKey = "NSPersistentHistoryTrackingKey"
            storeDescription.setOption(true as NSNumber, forKey: historyTrackingKey)

            container.persistentStoreDescriptions = [ storeDescription ]
            */
        }
        container.loadPersistentStores(completionHandler: { (storeDescription, error) in
            if let error = error as NSError? {
                // Replace this implementation with code to handle the error appropriately.
                // fatalError() causes the application to generate a crash log and terminate. You should not use this function in a shipping application, although it may be useful during development.

                /*
                Typical reasons for an error here include:
                * The parent directory does not exist, cannot be created, or disallows writing.
                * The persistent store is not accessible, due to permissions or data protection when the device is locked.
                * The device is out of space.
                * The store could not be migrated to the current model version.
                Check the error message to determine what the actual problem was.
                */
                fatalError("Unresolved error \(error), \(error.userInfo)")
            }
        })
        
        if !inMemory {
            
        }
    }
}
