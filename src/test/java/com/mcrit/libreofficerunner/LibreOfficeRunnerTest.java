/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mcrit.libreofficerunner;

import javax.json.JsonObject;
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
    }
    
    @After
    public void tearDown() {
    }

    /**
     * Test of recalculateXLXSFile method, of class LibreOfficeRunner.
     */
    @org.junit.Test
    public void testRecalculateXLXSFile() throws Exception {
        System.out.println("recalculateXLXSFile");
        String filePath = "";
        LibreOfficeRunner.recalculateXLXSFile(filePath);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of compileTemplate method, of class LibreOfficeRunner.
     */
    @org.junit.Test
    public void testCompileTemplate() {
        System.out.println("compileTemplate");
        String templateURL = "";
        JsonObject cellData = null;
        LibreOfficeRunner.compileTemplate(templateURL, cellData);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of main method, of class LibreOfficeRunner.
     */
    @org.junit.Test
    public void testMain() throws Exception {
        System.out.println("main");
        String[] args = null;
        LibreOfficeRunner.main(args);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }
    
}
