/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package bad.robot.excel.style;

import org.apache.poi.ss.usermodel.IndexedColors;

import static org.apache.poi.ss.usermodel.IndexedColors.*;

public enum Color {

    Brown(BROWN),
    Blue(LIGHT_BLUE),
    DarkRed(DARK_RED),
    DarkGrey(GREY_25_PERCENT),
    DarkYellow(YELLOW),
    Red(RED),
    Black(BLACK),
    Grey(GREY_25_PERCENT),
    White(WHITE),
    BrightGreen(BRIGHT_GREEN),
    Yellow(LIGHT_YELLOW),
    Pink(PINK),
    Turquoise(LIGHT_TURQUOISE),
    Green(LIGHT_GREEN),
    Violet(VIOLET),
    Teal(TEAL),
    Maroon(MAROON),
    Coral(CORAL),
    Rose(ROSE),
    Lavender(LAVENDER),
    Orange(LIGHT_ORANGE),
    Olive(OLIVE_GREEN),
    Plum(PLUM);

    private final IndexedColors color;

    Color(IndexedColors color) {
        this.color = color;
    }

    public short getPoiStyle() {
        return color.getIndex();
    }

}
