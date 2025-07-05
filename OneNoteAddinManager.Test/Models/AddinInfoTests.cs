using System;
using System.ComponentModel;
using System.Linq;
using OneNoteAddinManager.Lib.Models;

namespace OneNoteAddinManager.Test.Models;

[TestClass]
public class AddinInfoTests
{
    [TestMethod]
    public void Name_PropertyChanged_RaisesEvent()
    {
        // Arrange
        var addin = new AddinInfo();
        var eventRaised = false;
        var propertyName = string.Empty;
        
        addin.PropertyChanged += (sender, e) => 
        {
            eventRaised = true;
            propertyName = e.PropertyName;
        };

        // Act
        addin.Name = "TestAddin";

        // Assert
        Assert.IsTrue(eventRaised);
        Assert.AreEqual("Name", propertyName);
        Assert.AreEqual("TestAddin", addin.Name);
    }

    [TestMethod]
    public void FriendlyName_PropertyChanged_RaisesEvent()
    {
        // Arrange
        var addin = new AddinInfo();
        var eventRaised = false;
        var propertyName = string.Empty;
        
        addin.PropertyChanged += (sender, e) => 
        {
            eventRaised = true;
            propertyName = e.PropertyName;
        };

        // Act
        addin.FriendlyName = "Test Add-in";

        // Assert
        Assert.IsTrue(eventRaised);
        Assert.AreEqual("FriendlyName", propertyName);
        Assert.AreEqual("Test Add-in", addin.FriendlyName);
    }

    [TestMethod]
    public void Description_PropertyChanged_RaisesEvent()
    {
        // Arrange
        var addin = new AddinInfo();
        var eventRaised = false;
        var propertyName = string.Empty;
        
        addin.PropertyChanged += (sender, e) => 
        {
            eventRaised = true;
            propertyName = e.PropertyName;
        };

        // Act
        addin.Description = "Test description";

        // Assert
        Assert.IsTrue(eventRaised);
        Assert.AreEqual("Description", propertyName);
        Assert.AreEqual("Test description", addin.Description);
    }

    [TestMethod]
    public void DllPath_PropertyChanged_RaisesEvent()
    {
        // Arrange
        var addin = new AddinInfo();
        var eventRaised = false;
        var propertyName = string.Empty;
        
        addin.PropertyChanged += (sender, e) => 
        {
            eventRaised = true;
            propertyName = e.PropertyName;
        };

        // Act
        addin.DllPath = @"C:\test\test.dll";

        // Assert
        Assert.IsTrue(eventRaised);
        Assert.AreEqual("DllPath", propertyName);
        Assert.AreEqual(@"C:\test\test.dll", addin.DllPath);
    }

    [TestMethod]
    public void Guid_PropertyChanged_RaisesEvent()
    {
        // Arrange
        var addin = new AddinInfo();
        var eventRaised = false;
        var propertyName = string.Empty;
        
        addin.PropertyChanged += (sender, e) => 
        {
            eventRaised = true;
            propertyName = e.PropertyName;
        };

        // Act
        addin.Guid = "{12345678-1234-1234-1234-123456789ABC}";

        // Assert
        Assert.IsTrue(eventRaised);
        Assert.AreEqual("Guid", propertyName);
        Assert.AreEqual("{12345678-1234-1234-1234-123456789ABC}", addin.Guid);
    }

    [TestMethod]
    public void IsEnabled_PropertyChanged_RaisesEvent()
    {
        // Arrange
        var addin = new AddinInfo();
        var eventRaised = false;
        var propertyName = string.Empty;
        
        addin.PropertyChanged += (sender, e) => 
        {
            eventRaised = true;
            propertyName = e.PropertyName;
        };

        // Act
        addin.IsEnabled = true;

        // Assert
        Assert.IsTrue(eventRaised);
        Assert.AreEqual("IsEnabled", propertyName);
        Assert.IsTrue(addin.IsEnabled);
    }

    [TestMethod]
    public void LoadBehavior_PropertyChanged_RaisesEvent()
    {
        // Arrange
        var addin = new AddinInfo();
        var eventRaised = false;
        var propertyName = string.Empty;
        
        addin.PropertyChanged += (sender, e) => 
        {
            eventRaised = true;
            propertyName = e.PropertyName;
        };

        // Act
        addin.LoadBehavior = 3;

        // Assert
        Assert.IsTrue(eventRaised);
        Assert.AreEqual("LoadBehavior", propertyName);
        Assert.AreEqual(3, addin.LoadBehavior);
    }

    [TestMethod]
    public void Status_WhenEnabled_ReturnsEnabled()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.IsEnabled = true;

        // Assert
        Assert.AreEqual("Enabled", addin.Status);
    }

    [TestMethod]
    public void Status_WhenDisabled_ReturnsDisabled()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.IsEnabled = false;

        // Assert
        Assert.AreEqual("Disabled", addin.Status);
    }

    [TestMethod]
    public void OfficeAddinRegistryPath_WithName_ReturnsCorrectPath()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.Name = "TestAddin";

        // Assert
        Assert.AreEqual(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\TestAddin", 
                       addin.OfficeAddinRegistryPath);
    }

    [TestMethod]
    public void AppIdRegistryPath_WithGuid_ReturnsCorrectPath()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.Guid = "{12345678-1234-1234-1234-123456789ABC}";

        // Assert
        Assert.AreEqual(@"HKEY_CLASSES_ROOT\AppID\{12345678-1234-1234-1234-123456789ABC}", 
                       addin.AppIdRegistryPath);
    }

    [TestMethod]
    public void AppIdRegistryPath_WithoutGuid_ReturnsNotAvailable()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act & Assert
        Assert.AreEqual("Not Available", addin.AppIdRegistryPath);
    }

    [TestMethod]
    public void ClsidRegistryPath_WithGuid_ReturnsCorrectPath()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.Guid = "{12345678-1234-1234-1234-123456789ABC}";

        // Assert
        Assert.AreEqual(@"HKEY_CLASSES_ROOT\CLSID\{12345678-1234-1234-1234-123456789ABC}", 
                       addin.ClsidRegistryPath);
    }

    [TestMethod]
    public void ClsidRegistryPath_WithoutGuid_ReturnsNotAvailable()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act & Assert
        Assert.AreEqual("Not Available", addin.ClsidRegistryPath);
    }

    [TestMethod]
    public void ProgIdRegistryPath_WithName_ReturnsCorrectPath()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.Name = "TestAddin";

        // Assert
        Assert.AreEqual(@"HKEY_CLASSES_ROOT\TestAddin", addin.ProgIdRegistryPath);
    }

    [TestMethod]
    public void ComClassName_WithName_ReturnsCorrectClassName()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.Name = "TestAddin";

        // Assert
        Assert.AreEqual("TestAddin.AddIn", addin.ComClassName);
    }

    [TestMethod]
    public void LoadBehaviorExplanation_LoadBehavior0_ReturnsDisabled()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.LoadBehavior = 0;

        // Assert
        Assert.AreEqual("Disabled - Add-in is not loaded", addin.LoadBehaviorExplanation);
    }

    [TestMethod]
    public void LoadBehaviorExplanation_LoadBehavior1_ReturnsLoadedOnce()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.LoadBehavior = 1;

        // Assert
        Assert.AreEqual("Loaded once - Add-in is loaded only on demand", addin.LoadBehaviorExplanation);
    }

    [TestMethod]
    public void LoadBehaviorExplanation_LoadBehavior2_ReturnsLoadedAtStartup()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.LoadBehavior = 2;

        // Assert
        Assert.AreEqual("Loaded at startup - Add-in is loaded when the application starts", addin.LoadBehaviorExplanation);
    }

    [TestMethod]
    public void LoadBehaviorExplanation_LoadBehavior3_ReturnsLoadedAtStartupAndOnDemand()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.LoadBehavior = 3;

        // Assert
        Assert.AreEqual("Loaded at startup and on demand - Add-in is loaded at startup and remains loaded", addin.LoadBehaviorExplanation);
    }

    [TestMethod]
    public void LoadBehaviorExplanation_LoadBehavior8_ReturnsConnectedOnDemand()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.LoadBehavior = 8;

        // Assert
        Assert.AreEqual("Connected on demand - Add-in is loaded only when requested by the user", addin.LoadBehaviorExplanation);
    }

    [TestMethod]
    public void LoadBehaviorExplanation_LoadBehavior9_ReturnsConnectedAtStartup()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.LoadBehavior = 9;

        // Assert
        Assert.AreEqual("Connected at startup - Add-in is loaded at startup and connected", addin.LoadBehaviorExplanation);
    }

    [TestMethod]
    public void LoadBehaviorExplanation_LoadBehavior16_ReturnsConnectedWithFirstDocument()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.LoadBehavior = 16;

        // Assert
        Assert.AreEqual("Connected with first document - Add-in is loaded when the first document is opened", addin.LoadBehaviorExplanation);
    }

    [TestMethod]
    public void LoadBehaviorExplanation_UnknownBehavior_ReturnsUnknownWithValue()
    {
        // Arrange
        var addin = new AddinInfo();
        
        // Act
        addin.LoadBehavior = 999;

        // Assert
        Assert.AreEqual("Unknown behavior (999)", addin.LoadBehaviorExplanation);
    }

    [TestMethod]
    public void RegistryKeys_WithCompleteInformation_ReturnsAllKeys()
    {
        // Arrange
        var addin = new AddinInfo 
        {
            Name = "TestAddin",
            FriendlyName = "Test Add-in",
            Description = "Test description",
            DllPath = @"C:\test\test.dll",
            Guid = "{12345678-1234-1234-1234-123456789ABC}",
            LoadBehavior = 3
        };

        // Act
        var keys = addin.RegistryKeys;

        // Assert
        Assert.AreEqual(7, keys.Count);
        
        // Check Office Add-in Registration key
        var officeKey = keys.FirstOrDefault(k => k.Purpose == "Office Add-in Registration");
        Assert.IsNotNull(officeKey);
        Assert.AreEqual(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\OneNote\AddIns\TestAddin", officeKey.Path);
        
        // Check AppID key
        var appIdKey = keys.FirstOrDefault(k => k.Purpose == "AppID Registration");
        Assert.IsNotNull(appIdKey);
        Assert.AreEqual(@"HKEY_CLASSES_ROOT\AppID\{12345678-1234-1234-1234-123456789ABC}", appIdKey.Path);
        
        // Check CLSID key
        var clsidKey = keys.FirstOrDefault(k => k.Purpose == "CLSID Registration");
        Assert.IsNotNull(clsidKey);
        Assert.AreEqual(@"HKEY_CLASSES_ROOT\CLSID\{12345678-1234-1234-1234-123456789ABC}", clsidKey.Path);
        
        // Check ProgID key
        var progIdKey = keys.FirstOrDefault(k => k.Purpose == "ProgID Class Registration");
        Assert.IsNotNull(progIdKey);
        Assert.AreEqual(@"HKEY_CLASSES_ROOT\TestAddin", progIdKey.Path);
    }

    [TestMethod]
    public void RegistryKeys_WithoutGuid_ReturnsLimitedKeys()
    {
        // Arrange
        var addin = new AddinInfo 
        {
            Name = "TestAddin",
            FriendlyName = "Test Add-in",
            Description = "Test description",
            LoadBehavior = 3
        };

        // Act
        var keys = addin.RegistryKeys;

        // Assert
        Assert.AreEqual(2, keys.Count);
        
        // Should have Office Add-in Registration and ProgID Class Registration only
        var officeKey = keys.FirstOrDefault(k => k.Purpose == "Office Add-in Registration");
        Assert.IsNotNull(officeKey);
        
        var progIdKey = keys.FirstOrDefault(k => k.Purpose == "ProgID Class Registration");
        Assert.IsNotNull(progIdKey);
        
        // Should not have GUID-dependent keys
        Assert.IsFalse(keys.Any(k => k.Purpose == "AppID Registration"));
        Assert.IsFalse(keys.Any(k => k.Purpose == "CLSID Registration"));
    }
}