/*
 * Copyright (c) 2012, bad robot (london) ltd.
 *
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

package bad.robot.excel.matchers;

import org.junit.Test;

import static bad.robot.excel.WorkbookResource.getWorkbook;
import static bad.robot.excel.matchers.SheetNumberEqualityMatcher.hasSameNumberOfSheetsAs;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.not;

// Using Assert (from JUnit)
//java.lang.AssertionError:
//        Expected: <2> sheet(s)
//        got: <org.apache.poi.hssf.usermodel.HSSFWorkbook@3b6f0be8>

// Using MatcherAssert:
//java.lang.AssertionError:
//        Expected: <2> sheet(s)
//        but: got <1> sheet(s)
//        at org.hamcrest.MatcherAssert.assertThat(MatcherAssert.java:20)

public class SheetNumberEqualityMatcherTest {

    @Test
    public void sheetNumbersAreEqual() throws Exception {
        assertThat(getWorkbook("sheetNumbersAreEqual.xls"), hasSameNumberOfSheetsAs(getWorkbook("sheetNumbersAreEqual.xls")));
    }

    @Test
    public void sheetNumbersAreNotEqual() throws Exception {
        assertThat(getWorkbook("sheetNumbersAreEqual.xls"), not(hasSameNumberOfSheetsAs(getWorkbook("sheetNumbersAreNotEqual.xls"))));
    }
}