package com.mcrit.libreofficerunner;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Cristian Lorenzo i Martínez <cristian.lorenzo.martinez@gmail.com>
 */
import com.sun.star.bridge.XUnoUrlResolver;

import com.sun.star.beans.PropertyValue;

import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.frame.XDesktop;
import com.sun.star.io.IOException;

import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XComponent;

import com.sun.star.sheet.XCalculatable;
import com.sun.star.uno.Exception;
import com.sun.star.util.CloseVetoException;

import com.sun.star.util.XCloseable;

import com.sun.star.io.BufferSizeExceededException;
import com.sun.star.io.NotConnectedException;
import com.sun.star.io.XOutputStream;

import javax.json.Json;
import javax.json.JsonObject;
import javax.json.JsonObjectBuilder;
import javax.json.JsonWriter;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import java.io.ByteArrayOutputStream;
import java.io.*;
import java.lang.*;

/**
 * Class used to do all recalculation using LibreOffice UNO API
 * @author Cristian Lorenzo Martinez <cristian.lorenzo.martinez@gmail.com>
 */


public class LibreOfficeRunner {
    
    private static Object rInitialObject;
    private static XMultiComponentFactory xOfficeFactory;
    private static Object desktop;
    private static XDesktop xDesktop;
    private XComponent document;
    
    /**
     *
     * @param serviceURL Connection parameter to Headless LibreOffice instance.
     * @throws com.sun.star.uno.Exception
     */
    public LibreOfficeRunner(String serviceURL) throws Exception, java.lang.Exception {
        // create default local component context
        XComponentContext xLocalContext =
            com.sun.star.comp.helper.Bootstrap.createInitialComponentContext(null);

        // initial serviceManager
        XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();

        // create a urlresolver
        Object urlResolver  = xLocalServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", xLocalContext );

        // query for the XUnoUrlResolver interface
        XUnoUrlResolver xUrlResolver =
            (XUnoUrlResolver) UnoRuntime.queryInterface( XUnoUrlResolver.class, urlResolver );
        // Import the object
        LibreOfficeRunner.rInitialObject = xUrlResolver.resolve(serviceURL);

        // XComponentContext
        if( null == rInitialObject ) {
            throw new RuntimeException("Unable to get Initial context");
        } else {
            LibreOfficeRunner.xOfficeFactory = (XMultiComponentFactory) UnoRuntime.queryInterface(
                XMultiComponentFactory.class, rInitialObject);
                
            LibreOfficeRunner.desktop = xOfficeFactory.createInstanceWithContext(
                        "com.sun.star.frame.Desktop", xLocalContext);
            LibreOfficeRunner.xDesktop = (XDesktop)UnoRuntime.queryInterface(XDesktop.class, desktop);
        }
    }
    
    private static String getFilterName(String fileName) {
        Pattern regEx = Pattern.compile("\\.(\\w+$)");
        Matcher match = regEx.matcher(fileName);
        String filterName;
        if (match.find()) {
            switch (match.group(1)) {
                case "xlsx":
                    filterName = "Calc Office Open XML";
                    break;
                case "xls":
                    filterName = "MS Excel 97";
                    break;
                case "ods":
                    filterName = "calc8";
                    break;
                default:
                    throw new RuntimeException("Cannot match the provided filename with a valid LibreOffice Filter");
            }
        } else {
           throw new IllegalArgumentException("The filename provided doesn't has a file extension."); 
        }
        return filterName;
    }
    
    private class keyValue {
        public String Name;
        public Object Value;
    }
    
    private static PropertyValue[] createLoaderProperties(String fileName, keyValue[] additionalProperties) {
        
        PropertyValue[] loaderProperties;
        if (additionalProperties == null) {
          loaderProperties = new PropertyValue[1];  
        } else {
          loaderProperties = new PropertyValue[additionalProperties.length + 1];
        }
        
        loaderProperties[0] = new com.sun.star.beans.PropertyValue();
        loaderProperties[0].Name = "FilterName";
        loaderProperties[0].Value = getFilterName(fileName);
        
        for (int i = 1; i < loaderProperties.length; i += 1) {
            loaderProperties[i] = new com.sun.star.beans.PropertyValue();
            loaderProperties[i].Name = additionalProperties[i - 1].Name;
            loaderProperties[i].Value = additionalProperties[i - 1].Value;
        }
        
        return loaderProperties;
    }
    
    private void streamDocumentToStdout(String fileExtension) throws IOException, CloseVetoException {
       XStorable xStorable;
       
       xStorable = (XStorable)UnoRuntime.queryInterface(
                XStorable.class, document);
       
       PropertyValue[] saveProperties = new PropertyValue[3];
       saveProperties[0] = new com.sun.star.beans.PropertyValue();
       saveProperties[0].Name = "FilterName";
       saveProperties[0].Value = getFilterName(fileExtension);
       
       saveProperties[1] = new com.sun.star.beans.PropertyValue();
       saveProperties[1].Name = "Overwrite";
       saveProperties[1].Value = true;
       
       saveProperties[2] = new com.sun.star.beans.PropertyValue();
       saveProperties[2].Name = "OutputStream";
       saveProperties[2].Value = new StdoutStream();
       
       xStorable.storeToURL("private:stream", saveProperties);
    }
    
    private static class StdoutStream extends ByteArrayOutputStream implements XOutputStream {

    private StdoutStream() {
        super(32768);
    }

    //
    // Implement XOutputStream
    //
    public void writeBytes(byte[] values) throws NotConnectedException, BufferSizeExceededException, com.sun.star.io.IOException {
        try {
            System.out.write(values);
        }
        catch (java.io.IOException e) {
            throw(new com.sun.star.io.IOException(e.getMessage()));
        }
    }

    public void closeOutput() throws NotConnectedException, BufferSizeExceededException, com.sun.star.io.IOException {
        try {
            super.flush();
            super.close();
        }
        catch (java.io.IOException e) {
            throw(new com.sun.star.io.IOException(e.getMessage()));
        }
    }

    @Override
    public void flush() {
        try {
            super.flush();
        }
        catch (java.io.IOException e) {
        }
    }
}
    
    private void closeDocument() throws CloseVetoException {
        XCloseable xCloseable;
        
        xCloseable = (XCloseable)UnoRuntime.queryInterface(
                            XCloseable.class, document);
        if ( xCloseable != null ) {
            xCloseable.close(false);
        } else {
            document.dispose();
        }
    }
    
    private void loadFileFromURL(String filePath, keyValue[] additionalProps) throws IOException, com.sun.star.lang.IllegalArgumentException {
        XComponentLoader xComponentLoader;      
        xComponentLoader = (XComponentLoader)UnoRuntime.queryInterface(
                XComponentLoader.class, LibreOfficeRunner.desktop);
        
        PropertyValue[] propertiesLoader = createLoaderProperties(filePath, additionalProps);
        
        document = xComponentLoader.loadComponentFromURL(
        "file://" + filePath, "_blank", 0, propertiesLoader);
    }
    
    private void overwriteFile() throws IOException {
        XStorable xStorable;
        
        xStorable = (XStorable)UnoRuntime.queryInterface(
                XStorable.class, document);

        xStorable.store();
    }
    
    private void recalculateAll() {
        XCalculatable xCalculatable;
        
        xCalculatable = (XCalculatable) UnoRuntime.queryInterface(
                XCalculatable.class, document);
        xCalculatable.calculateAll();
    }
    
    /**
     *
     * @param filePath
     * @throws IOException
     * @throws CloseVetoException
     * @throws com.sun.star.lang.IllegalArgumentException
     */
    public void recalculateXLXSFile(String filePath) throws IOException, CloseVetoException, com.sun.star.lang.IllegalArgumentException {
     
        keyValue[] propertiesLoader = new keyValue[1];
        propertiesLoader[0].Name = "Overwrite";
        propertiesLoader[0].Value = true; 
        loadFileFromURL(filePath, propertiesLoader);
        recalculateAll();
        overwriteFile();
        closeDocument();              
    }
    
    /**
     *
     * @param templateURL
     * @param cellData
     */
    public static void compileTemplate(String templateURL, JsonObject cellData ) {
        
    }
    
    /**
     *
     * @param args
     * @throws java.lang.Exception
     */
    
    
    
    public static void main(String[] args) throws java.lang.Exception {
        LibreOfficeRunner instance;
        instance = new LibreOfficeRunner("uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager");
        
        keyValue[] properties = new keyValue[0];
        instance.loadFileFromURL("/home/mcrituser/test.ods", properties);
        System.setOut(new PrintStream(new BufferedOutputStream(new FileOutputStream("test.xls")), true));
        instance.streamDocumentToStdout(".xls");
        instance.closeDocument();
        System.exit(0);
    }
}