/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mcrit.libreofficerunner;

import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.io.StringReader;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.json.Json;
import javax.json.JsonReader;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author mcrituser
 */
public class LibreOfficeRunnerTest {
    private final ByteArrayOutputStream outContent = new ByteArrayOutputStream(); 
    
    public LibreOfficeRunnerTest() {
        
    }
    
    @BeforeClass
    public static void setUpClass() { 
        
    }
    
    @AfterClass
    public static void tearDownClass() {
    }
    
    @Before
    public void setUp() {
        System.setOut(new PrintStream(outContent));
    }
    
    @After
    public void tearDown() {
        System.setOut(null);
    }

    @Test
    public void compileTemplate() throws FileNotFoundException {
         
    }

    /**
     * Test of recalculateFile method, of class LibreOfficeRunner.
     */
    /*@Test
    public void testRecalculateFile() throws Exception {
        System.out.println("recalculateFile");
        String filePath = "";
        LibreOfficeRunner instance = null;
        instance.recalculateFile(filePath);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }*/

    /**
     * Test of compileTemplate method, of class LibreOfficeRunner.
     * @throws java.lang.Exception
     */
    @Test
    public void testCompileTemplate() throws Exception {
        LibreOfficeRunner instance = new LibreOfficeRunner("uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager");
        
        JsonReader jsonReader = Json.createReader(new StringReader("[ {\"target\" : [ 0, [0, 0]], \"data\": [[1, 2], [3,4]]}]"));

        instance.compileTemplate("/home/mcrituser/test.ods", ".csv", jsonReader.readArray()); 
        assertEquals("The compiled template is doen't match the given JSON input.", "1,2\n3,4\n", outContent.toString());
    }
}
