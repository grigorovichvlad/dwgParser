#include <windows.h>
#include <tchar.h>
#include <comdef.h>
#include <acadi.h>
#include <locale>
#import "C:\ObjectARX\inc-x64\acax24ENU.tlb" // Path to acax24enu.tlb
#include <atlbase.h>
#include <atlcom.h>
#include <iostream>
#include <fstream>
#include <string>
#include <map>
#include <dbmain.h>
#include <set>
#include <dbents.h>
#include <dbobjptr.h>
#include <aced.h>
#include <dbtable.h>
#include <sstream>
#include <vector>
#include <acdb.h>
#include <dbapserv.h>
#include <dbmain.h>
#include <dbents.h>
#include <dbapserv.h>

std::map<std::wstring, IAcadEntity*> entitiesMap;
std::set<std::wstring> entitiesNames;

std::wstring handleEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity, std::wstring formattedEntityName);

const std::wstring outputFile = L"D:\output.txt";

void DebugPrint(const TCHAR* message) {
    OutputDebugString(message);
    OutputDebugString(_T("\n"));
}

std::wstring getFormattedEntityName(BSTR entityName) {
    std::wstring name = (wchar_t*)entityName;
    std::wstring prefix = L"AcDb";

    //// Check if the name starts with "AcDb"
    if (name.substr(0, prefix.size()) == prefix) {
        // If it does, remove "AcDb" from the name
        name = name.substr(prefix.size());
    }

    // Convert the name to lowercase
    std::transform(name.begin(), name.end(), name.begin(), ::tolower);

    return name;
}

void saveStringToFile(const std::wstring& str, const std::wstring& filename) {
    std::wofstream outputFile(filename, std::ios::app);
    if (outputFile.is_open()) {
        outputFile << str;
        outputFile.close();
        //std::wcout << L"Строка успешно сохранена в файл: " << filename << std::endl;
    }
    else {
        std::wcerr << L"Ошибка при открытии файла: " << filename << std::endl;
    }
}

std::wstring handleLineEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadLine* line;  // Line interface pointer
    // Query for the IAcadLine interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadLine), (void**)&line);
    if (SUCCEEDED(hr))
    {
        // Initialize variants
        VARIANT startVar;
        VariantInit(&startVar);
        VARIANT endVar;
        VariantInit(&endVar);

        // Get points
        line->get_StartPoint(&startVar);
        line->get_EndPoint(&endVar);

        // Assume that the returned variants are arrays of doubles
        if (startVar.vt == (VT_ARRAY | VT_R8) && endVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saStart = startVar.parray;
            SAFEARRAY* saEnd = endVar.parray;

            double* startArr;
            double* endArr;

            SafeArrayAccessData(saStart, (void**)&startArr);
            SafeArrayAccessData(saEnd, (void**)&endArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second << L" Line from (" << startArr[0] << L", " << startArr[1] << L", " << startArr[2] << L") to ("
                << endArr[0] << L", " << endArr[1] << L", " << endArr[2] << L")" << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saStart);
            SafeArrayUnaccessData(saEnd);
        }

        // Clear the variants
        VariantClear(&startVar);
        VariantClear(&endVar);
    }
    return output;
}

std::wstring handleArcEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadArc* arc;  // Arc interface pointer
    // Query for the IAcadArc interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadArc), (void**)&arc);
    if (SUCCEEDED(hr))
    {
        // Initialize variants
        VARIANT startVar;
        VariantInit(&startVar);
        VARIANT endVar;
        VariantInit(&endVar);
        VARIANT centerVar;
        VariantInit(&centerVar);
        double radius;
        double startAngle;
        double endAngle;

        // Get points and parameters
        arc->get_StartPoint(&startVar);
        arc->get_EndPoint(&endVar);
        arc->get_Center(&centerVar);
        arc->get_Radius(&radius);
        arc->get_StartAngle(&startAngle);
        arc->get_EndAngle(&endAngle);

        // Assume that the returned variants are arrays of doubles
        if (startVar.vt == (VT_ARRAY | VT_R8) && endVar.vt == (VT_ARRAY | VT_R8) && centerVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saStart = startVar.parray;
            SAFEARRAY* saEnd = endVar.parray;
            SAFEARRAY* saCenter = centerVar.parray;

            double* startArr;
            double* endArr;
            double* centerArr;

            SafeArrayAccessData(saStart, (void**)&startArr);
            SafeArrayAccessData(saEnd, (void**)&endArr);
            SafeArrayAccessData(saCenter, (void**)&centerArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << L": " << entity.second << L" Arc from (" << startArr[0] << L", " << startArr[1] << L", " << startArr[2] << L") to ("
                << endArr[0] << L", " << endArr[1] << L", " << endArr[2] << L") Center at ("
                << centerArr[0] << L", " << centerArr[1] << L", " << centerArr[2] << L") Radius: " << radius
                << L" StartAngle: " << startAngle << L" EndAngle: " << endAngle << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saStart);
            SafeArrayUnaccessData(saEnd);
            SafeArrayUnaccessData(saCenter);
        }

        // Clear the variants
        VariantClear(&startVar);
        VariantClear(&endVar);
        VariantClear(&centerVar);
    }
    return output;
}

std::wstring handleAttributeEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadAttribute* attribute;  // Attribute interface pointer
    // Query for the IAcadAttribute interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadAttribute), (void**)&attribute);
    if (SUCCEEDED(hr))
    {
        // Initialize BSTRs for the properties
        BSTR bstrTag, bstrPrompt, bstrTextString;

        // Get attribute data
        attribute->get_TagString(&bstrTag);
        attribute->get_PromptString(&bstrPrompt);
        attribute->get_TextString(&bstrTextString);

        std::wcout << L"ID " << entity.first << L": " << entity.second << L" Attribute: Tag - " << bstrTag << L", Prompt - " << bstrPrompt << L", TextString - " << bstrTextString << std::endl;

        // Clean up
        SysFreeString(bstrTag);
        SysFreeString(bstrPrompt);
        SysFreeString(bstrTextString);
    }
    return output;
}

std::wstring handleBlockReferenceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadBlockReference* blockRef;  // Block reference interface pointer

    // Query for the IAcadBlockReference interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadBlockReference), (void**)&blockRef);

    if (SUCCEEDED(hr)) {
        // Explode the block reference
        VARIANT explodedEntitiesVar;
        VariantInit(&explodedEntitiesVar);

        hr = blockRef->Explode(&explodedEntitiesVar);
        if (SUCCEEDED(hr) && explodedEntitiesVar.vt == (VT_ARRAY | VT_DISPATCH)) {
            SAFEARRAY* saExplodedEntities = explodedEntitiesVar.parray;
            LONG lowerBound, upperBound;
            SafeArrayGetLBound(saExplodedEntities, 1, &lowerBound);
            SafeArrayGetUBound(saExplodedEntities, 1, &upperBound);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << L": " << entity.second << L" Block Reference exploded entities:" << std::endl;

            for (LONG i = lowerBound; i <= upperBound; i++) {
                IDispatch* dispEntity;
                SafeArrayGetElement(saExplodedEntities, &i, &dispEntity);

                // Cast the IDispatch pointer to the IAcadEntity interface
                IAcadEntity* entity;
                hr = dispEntity->QueryInterface(__uuidof(IAcadEntity), (void**)&entity);
                if (SUCCEEDED(hr)) {
                    // Now you can interact with the individual entities that comprised the block
                    // Here we're getting the handle of each entity
                    BSTR handle;
                    entity->get_Handle(&handle);
                    BSTR entityName;
                    entity->get_EntityName(&entityName);
                    std::wstring formattedEntityName = getFormattedEntityName(entityName);
                    std::wcout << L"Entity ID: " << handle << " Entity name: " << formattedEntityName << std::endl;

                    if (formattedEntityName == L"BlockReference") {
                        // Ask the user if they want to explore this entity
                        std::wcout << L"Do you want to explore this block? (y/n): ";
                        char choice;
                        std::cin >> choice;

                        if (tolower(choice) == 'y') {
                            // Process each entity recursively
                            handleEntity({ std::wstring(handle, SysStringLen(handle)), entity }, formattedEntityName);
                        }
                    }

                    // Free the handle
                    SysFreeString(handle);
                }

                // Release the entity
                entity->Release();
                dispEntity->Release();
            }

            std::wcout.rdbuf(oldWcoutStreamBuf);
            output = oss.str();

            std::wcout << output << std::endl;

            // Destroy the safe array
            SafeArrayDestroy(saExplodedEntities);
        }
    }

    return output;
}

std::wstring handleBlock(const std::pair<std::wstring, CComPtr<IAcadEntity>>& block) {
    std::wstring output;
    IAcadBlock* blockPtr;  // Block interface pointer

    // Query for the IAcadBlock interface from the block
    HRESULT hr = block.second->QueryInterface(__uuidof(IAcadBlock), (void**)&blockPtr);

    if (SUCCEEDED(hr)) {
        BSTR handle;
        blockPtr->get_Handle(&handle);

        BSTR name;
        blockPtr->get_Name(&name);

        VARIANT_BOOL isXRef;
        blockPtr->get_IsXRef(&isXRef);

        VARIANT_BOOL isLayout;
        blockPtr->get_IsLayout(&isLayout);

        VARIANT_BOOL isDynamicBlock;
        blockPtr->get_IsDynamicBlock(&isDynamicBlock);

        std::wostringstream oss;
        std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

        std::wcout << L"Block ID: " << handle
            << L" Name: " << name
            << L" IsXRef: " << ((isXRef == VARIANT_TRUE) ? "Yes" : "No")
            << L" IsLayout: " << ((isLayout == VARIANT_TRUE) ? "Yes" : "No")
            << L" IsDynamicBlock: " << ((isDynamicBlock == VARIANT_TRUE) ? "Yes" : "No")
            << std::endl;

        std::wcout.rdbuf(oldWcoutStreamBuf);
        output = oss.str();

        // Free the handles
        SysFreeString(handle);
        SysFreeString(name);
    }

    return output;
}

std::wstring handleCircleEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadCircle* circle;  // Circle interface pointer
    // Query for the IAcadCircle interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadCircle), (void**)&circle);
    if (SUCCEEDED(hr))
    {
        if (circle != nullptr) {
            // Initialize variants
            VARIANT centerVar;
            VariantInit(&centerVar);
            double radius;

            // Get center point and radius
            circle->get_Center(&centerVar);
            circle->get_Radius(&radius);

            // Assume that the returned variant is an array of doubles
            if (centerVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCenter = centerVar.parray;

                double* centerArr;

                SafeArrayAccessData(saCenter, (void**)&centerArr);

                std::wcout << L"ID " << entity.first << L": " << entity.second << L"; Circle with center (" << centerArr[0] << L", " << centerArr[1] << L", " << centerArr[2] << L") and radius "
                    << radius << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCenter);
            }

            // Clear the variant
            VariantClear(&centerVar);
        }
    }
    return output;
}

std::wstring handleEllipseEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadEllipse* ellipse;  // Ellipse interface pointer
    // Query for the IAcadEllipse interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadEllipse), (void**)&ellipse);
    if (SUCCEEDED(hr))
    {
        if (ellipse != nullptr) {
            // Initialize variants
            VARIANT centerVar;
            VariantInit(&centerVar);
            double majorAxisVar;
            double minorAxisVar;

            // Get ellipse details
            ellipse->get_Center(&centerVar);
            ellipse->get_MajorRadius(&majorAxisVar);
            ellipse->get_MinorRadius(&minorAxisVar);

            // Assume that the returned center variant is an array of doubles
            // And the returned major and minor axis variants are doubles
            if (centerVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCenter = centerVar.parray;

                double* centerArr;

                SafeArrayAccessData(saCenter, (void**)&centerArr);

                std::wcout << L"ID " << entity.first << L": " << entity.second << L" Ellipse at center (" << centerArr[0] << L", " << centerArr[1] << L", " << centerArr[2] << L") with major axis "
                    << majorAxisVar << L" and minor axis " << minorAxisVar << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCenter);
            }

            // Clear the variants
            VariantClear(&centerVar);
        }
    }
    return output;
}

std::wstring handleExtrudedSurfaceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadExtrudedSurface* extrudedSurface;  // ExtrudedSurface interface pointer
    // Query for the IAcadExtrudedSurface interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadExtrudedSurface), (void**)&extrudedSurface);
    if (SUCCEEDED(hr))
    {
        if (extrudedSurface != nullptr) {
            double heightVar;

            // Get profile and height
            extrudedSurface->get_Height(&heightVar);
            std::wcout << L"ID " << entity.first << L": " << entity.second << L" and Height: " << heightVar << std::endl;
        }
    }
    return output;
}

std::wstring handleHatchEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadHatch* hatch;  // Hatch interface pointer
    // Query for the IAcadHatch interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadHatch), (void**)&hatch);
    if (SUCCEEDED(hr))
    {
        if (hatch != nullptr) {
            // Initialize variants
            double areaVar;
            BSTR patternNameVar;

            // Get properties
            hatch->get_Area(&areaVar);
            hatch->get_PatternName(&patternNameVar);


            std::wcout << L"ID " << entity.first << L": " << entity.second << L" Hatch area (" << areaVar << ") with pattern " << (wchar_t*)patternNameVar << std::endl;

        }
    }
    return output;
}

std::wstring handleHelixEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadHelix* helix;  // Helix interface pointer
    // Query for the IAcadHelix interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadHelix), (void**)&helix);
    if (SUCCEEDED(hr))
    {
        if (helix != nullptr) {
            // Initialize variants
            VARIANT basePointVar;
            double heightVar, radiusVar;
            ACAD_NOUNITS turnsVar;
            VariantInit(&basePointVar);

            // Get properties
            helix->get_Position(&basePointVar);
            helix->get_BaseRadius(&radiusVar);
            helix->get_Turns(&turnsVar);
            helix->get_Height(&heightVar);

            // Assume that the returned variants are arrays of doubles for points
            // and doubles for other properties
            if (basePointVar.vt == (VT_ARRAY | VT_R8)
                && turnsVar == VT_R8 && heightVar == VT_R8) {

                SAFEARRAY* saBase = basePointVar.parray;

                double* baseArr;

                SafeArrayAccessData(saBase, (void**)&baseArr);

                std::wcout << L"ID " << entity.first << L": " << entity.second << L" Helix base point ("
                    << baseArr[0] << L", " << baseArr[1] << L", " << baseArr[2] << L") with radius "
                    << radiusVar << L", turns " << turnsVar << L", and height " << heightVar << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saBase);
            }

            // Clear the variants
            VariantClear(&basePointVar);
        }
    }
    return output;
}

std::wstring handleLeaderEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadLeader* leader;  // Leader interface pointer
    // Query for the IAcadLeader interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadLeader), (void**)&leader);
    if (SUCCEEDED(hr))
    {
        if (leader != nullptr) {
            // Initialize variants
            VARIANT coordinatesVar;
            VariantInit(&coordinatesVar);

            double textGapVar;

            // Get coordinates and text gap
            leader->get_Coordinates(&coordinatesVar);
            leader->get_TextGap(&textGapVar);

            // Assume that the returned variant is an array of doubles for coordinates and a double for TextGap
            if (coordinatesVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoordinates = coordinatesVar.parray;

                double* coordinatesArr;
                long lBound, uBound;
                SafeArrayGetLBound(saCoordinates, 1, &lBound);
                SafeArrayGetUBound(saCoordinates, 1, &uBound);
                SafeArrayAccessData(saCoordinates, (void**)&coordinatesArr);

                for (long i = lBound; i <= uBound; i += 3) {
                    std::wcout << L"ID " << entity.first << L": " << entity.second << L" Leader with coordinates at (" << coordinatesArr[i] << L", " << coordinatesArr[i + 1] << L", " << coordinatesArr[i + 2] << L")" << std::endl;
                }

                std::wcout << L"Text gap is " << textGapVar << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoordinates);
            }

            // Clear the variants
            VariantClear(&coordinatesVar);
        }
    }
    return output;
}

std::wstring handleLoftedSurfaceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadLoftedSurface* loftedSurface;  // LoftedSurface interface pointer
    // Query for the IAcadLoftedSurface interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadLoftedSurface), (void**)&loftedSurface);
    if (SUCCEEDED(hr))
    {
        if (loftedSurface != nullptr)
        {
            double startDraftAngle;
            double endDraftAngle;
            double startSmoothMagnitude;
            double endSmoothMagnitude;

            loftedSurface->get_StartDraftAngle(&startDraftAngle);
            loftedSurface->get_EndDraftAngle(&endDraftAngle);
            loftedSurface->get_StartSmoothMagnitude(&startSmoothMagnitude);
            loftedSurface->get_EndSmoothMagnitude(&endSmoothMagnitude);

            std::wcout << L"ID " << entity.first << L": " << entity.second << L" LoftedSurface with StartDraftAngle: " << startDraftAngle
                << L", EndDraftAngle: " << endDraftAngle
                << L", StartSmoothMagnitude: " << startSmoothMagnitude
                << L", EndSmoothMagnitude: " << endSmoothMagnitude << std::endl;
        }
    }
    return output;
}

std::wstring handleMInsertBlockEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadMInsertBlock* block;  // Block interface pointer
    // Query for the IAcadMInsertBlock interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadMInsertBlock), (void**)&block);
    if (SUCCEEDED(hr))
    {
        if (block != nullptr) {
            // Initialize variants
            VARIANT insertionPoint;
            VariantInit(&insertionPoint);

            // Get points
            block->get_InsertionPoint(&insertionPoint);

            // Assume that the returned variant is an array of doubles
            if (insertionPoint.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saInsertion = insertionPoint.parray;

                double* insertionArr;

                SafeArrayAccessData(saInsertion, (void**)&insertionArr);

                std::wcout << L"ID " << entity.first << L": " << entity.second << L" Block at (" << insertionArr[0] << L", " << insertionArr[1] << L", " << insertionArr[2] << L")" << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saInsertion);
            }

            // Get Rows
            long rows;
            block->get_Rows(&rows);

            std::wcout << L"Rows: " << rows << std::endl;

            // Get Columns
            long columns;
            block->get_Columns(&columns);

            std::wcout << L"Columns: " << columns << std::endl;

            // Get Name
            BSTR name;
            block->get_Name(&name);

            std::wcout << L"Name: " << name << std::endl;

            // Get XScaleFactor
            double xscale;
            block->get_XScaleFactor(&xscale);

            std::wcout << L"XScaleFactor: " << xscale << std::endl;

            // Clear the variant
            VariantClear(&insertionPoint);
        }
    }
    return output;
}

std::wstring handleMLeaderEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadMLeader* mleader;  // MLeader interface pointer
    // Query for the IAcadMLeader interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadMLeader), (void**)&mleader);
    if (SUCCEEDED(hr))
    {
        if (mleader != nullptr) {
            // Initialize variant
            BSTR textVar;

            // Get text string
            mleader->get_TextString(&textVar);

            BSTR textBstr = textVar;
            std::wcout << L"ID " << entity.first << L": " << entity.second << L" MLeader with text: " << textBstr << std::endl;

            // Get ContentType
            AcMLeaderContentType contentType;
            mleader->get_ContentType(&contentType);
            std::wcout << L"ID " << entity.first << L": " << entity.second << L" MLeader with content type: " << contentType << std::endl;
        }
    }
    return output;
}

std::wstring handleMLineEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadMLine* mLine;  // MLine interface pointer
    // Query for the IAcadMLine interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadMLine), (void**)&mLine);
    if (SUCCEEDED(hr))
    {
        if (mLine != nullptr) {
            // Initialize variants
            VARIANT coordinatesVar;
            VariantInit(&coordinatesVar);

            AcMLineJustification justificationVar;

            double mLineScaleVar;

            BSTR styleNameVar;

            // Get properties
            mLine->get_Coordinates(&coordinatesVar);
            mLine->get_Justification(&justificationVar);
            mLine->get_MLineScale(&mLineScaleVar);
            mLine->get_StyleName(&styleNameVar);

            // Assume that the returned coordinates variant is an array of doubles
            if (coordinatesVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoordinates = coordinatesVar.parray;

                double* coordinatesArr;

                SafeArrayAccessData(saCoordinates, (void**)&coordinatesArr);

                std::wcout << L"ID " << entity.first << L": " << entity.second << L" MLine with coordinates ("
                    << coordinatesArr[0] << L", " << coordinatesArr[1] << L", " << coordinatesArr[2] << L") and " << L"Justification: " << justificationVar <<
                    L"StyleName: " << styleNameVar << L"MLineScale: " << mLineScaleVar << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoordinates);
            }

            // Clear the variants
            VariantClear(&coordinatesVar);
        }
    }
    return output;
}

std::wstring handleMTextEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadMText* mtext;  // MText interface pointer
    // Query for the IAcadMText interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadMText), (void**)&mtext);
    if (SUCCEEDED(hr))
    {
        if (mtext != nullptr) {
            // Initialize variants
            VARIANT insertionPointVar;
            VariantInit(&insertionPointVar);
            BSTR textStringVar;
            double heightVar;
            ACAD_ANGLE rotationVar;
            BSTR styleNameVar;

            // Get properties
            mtext->get_InsertionPoint(&insertionPointVar);
            mtext->get_TextString(&textStringVar);
            mtext->get_Height(&heightVar);
            mtext->get_Rotation(&rotationVar);
            mtext->get_StyleName(&styleNameVar);

            // Assume that the returned variants are of the appropriate types
            if (insertionPointVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saInsertionPoint = insertionPointVar.parray;
                BSTR textString = textStringVar;
                double height = heightVar;
                double rotation = rotationVar;
                BSTR styleName = styleNameVar;

                double* insertionPointArr;
                SafeArrayAccessData(saInsertionPoint, (void**)&insertionPointArr);

                std::wcout << L"ID " << entity.first << L": " << entity.second << std::endl;
                std::wcout << L"Insertion Point: (" << insertionPointArr[0] << L", " << insertionPointArr[1] << L", " << insertionPointArr[2] << L")" << std::endl;
                std::wcout << L"Text String: " << textString << std::endl;
                std::wcout << L"Height: " << height << std::endl;
                std::wcout << L"Rotation: " << rotation << std::endl;
                std::wcout << L"Style Name: " << styleName << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saInsertionPoint);
            }

            // Clear the variants
            VariantClear(&insertionPointVar);
        }
    }
    return output;
}

std::wstring handleNurbsSurfaceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadNurbSurface* nurbsSurface;  // NurbSurface interface pointer            
    // Query for the IAcadNurbSurface interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadNurbSurface), (void**)&nurbsSurface);

    if (SUCCEEDED(hr))
    {
        if (nurbsSurface != nullptr) {
            // Get some properties
            BSTR surfaceTypeVar;
            nurbsSurface->get_SurfaceType(&surfaceTypeVar);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << L": " << entity.second << L" NURBS Surface with Surface Type " << surfaceTypeVar << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;
        }
    }
    return output;
}

std::wstring handleOleObjectEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadOle* oleObj;  // IAcadOle interface pointer
    // Query for the IAcadOle interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadOle), (void**)&oleObj);
    if (SUCCEEDED(hr))
    {
        if (oleObj != nullptr) {
            // Initialize variants for each property
            VARIANT oleInsertionPoint;
            VARIANT_BOOL  oleVisible;
            double oleHeight, oleWidth, oleScaleHeight, oleScaleWidth;
            BSTR oleEntityTransparency, oleOleSourceApp;
            AcOleType oleOleItemType;
            ACAD_ANGLE oleRotation;

            VariantInit(&oleInsertionPoint);

            // Get properties
            oleObj->get_Height(&oleHeight);
            oleObj->get_InsertionPoint(&oleInsertionPoint);
            oleObj->get_Width(&oleWidth);
            oleObj->get_EntityTransparency(&oleEntityTransparency);
            oleObj->get_OleItemType(&oleOleItemType);
            oleObj->get_OleSourceApp(&oleOleSourceApp);
            oleObj->get_Rotation(&oleRotation);
            oleObj->get_ScaleHeight(&oleScaleHeight);
            oleObj->get_ScaleWidth(&oleScaleWidth);
            oleObj->get_Visible(&oleVisible);

            // Print the properties
            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << L": " << entity.second << L" OleObject with Height: " << oleHeight
                << L", Width: " << oleWidth << L", Insertion Point: (" << oleInsertionPoint.dblVal << L"), Entity Transparency: "
                << oleEntityTransparency << L", Ole Item Type: " << oleOleItemType << L", Ole Source App: "
                << oleOleSourceApp << L", Rotation: " << oleRotation << L", Scale Height: " << oleScaleHeight
                << L", Scale Width: " << oleScaleWidth << L", Visible: " << (oleVisible ? L"True" : L"False") << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Clear the variants
            VariantClear(&oleInsertionPoint);
        }
    }
    return output;
}

std::wstring handlePlaneSurfaceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadPlaneSurface* planeSurface;  // PlaneSurface interface pointer
    // Query for the IAcadPlaneSurface interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadPlaneSurface), (void**)&planeSurface);
    if (SUCCEEDED(hr))
    {
        if (planeSurface != nullptr) {
            // Initialize variants
            BSTR surfaceTypeVar, materialVar, lineTypeVar;

            // Get properties
            planeSurface->get_SurfaceType(&surfaceTypeVar);
            planeSurface->get_Material(&materialVar);
            planeSurface->get_Linetype(&lineTypeVar);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << L": " << entity.second
                << L" PlaneSurface with SurfaceType " << surfaceTypeVar
                << L", Material " << materialVar
                << L", and LineType " << lineTypeVar << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;
        }
    }
    return output;
}

std::wstring handlePointEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadPoint* point;  // Point interface pointer
    // Query for the IAcadPoint interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadPoint), (void**)&point);
    if (SUCCEEDED(hr))
    {
        if (point != nullptr) {
            // Initialize variant
            VARIANT coordVar;
            VariantInit(&coordVar);

            // Get point coordinates
            point->get_Coordinates(&coordVar);

            // Assume that the returned variant is an array of doubles
            if (coordVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoords = coordVar.parray;

                double* coordArr;

                SafeArrayAccessData(saCoords, (void**)&coordArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << L": " << entity.second << L" Point at (" << coordArr[0] << L", " << coordArr[1] << L", " << coordArr[2] << L")" << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoords);
            }

            // Clear the variant
            VariantClear(&coordVar);
        }
    }
    return output;
}

std::wstring handlePointCloudEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadPointCloud* pointCloud;  // PointCloud interface pointer
    // Query for the IAcadPointCloud interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadPointCloud), (void**)&pointCloud);
    if (SUCCEEDED(hr))
    {
        if (pointCloud != nullptr) {
            // Initialize variants
            BSTR nameVar;
            ACAD_DISTANCE heightVar, lengthVar, widthVar;
            VARIANT insertionPointVar;
            VariantInit(&insertionPointVar);

            // Get properties
            pointCloud->get_Name(&nameVar);
            pointCloud->get_Height(&heightVar);
            pointCloud->get_Length(&lengthVar);
            pointCloud->get_Width(&widthVar);
            pointCloud->get_InsertionPoint(&insertionPointVar);

            // Assume that the returned variants are arrays of doubles for insertion point
            if (insertionPointVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saInsertionPoint = insertionPointVar.parray;
                double* insertionPointArr;
                SafeArrayAccessData(saInsertionPoint, (void**)&insertionPointArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" PointCloud name: " << nameVar << L", Height: " << heightVar << L", Length: " << lengthVar
                    << L", Width: " << widthVar << L", InsertionPoint: (" << insertionPointArr[0] << L", " << insertionPointArr[1] << L", " << insertionPointArr[2] << L")" << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saInsertionPoint);
            }

            // Clear the variants
            VariantClear(&insertionPointVar);
        }
    }
    return output;
}

std::wstring handlePointCloudExEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadPointCloudEx* pointCloudEx;  // PointCloud interface pointer
    // Query for the IAcadPointCloudEx interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadPointCloudEx), (void**)&pointCloudEx);

    if (SUCCEEDED(hr))
    {
        if (pointCloudEx != nullptr) {
            // Initialize variants
            BSTR name;
            VARIANT insertionPointVar;
            VariantInit(&insertionPointVar);

            // Get name and insertion point
            pointCloudEx->get_Name(&name);
            pointCloudEx->get_InsertionPoint(&insertionPointVar);

            // Assume that the returned variants are BSTR and arrays of doubles
            if (insertionPointVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saInsertion = insertionPointVar.parray;

                double* insertionArr;

                SafeArrayAccessData(saInsertion, (void**)&insertionArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" PointCloud '" << name << L"' with insertion point at (" << insertionArr[0] << L", " << insertionArr[1] << L", " << insertionArr[2] << L")" << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saInsertion);
            }

            // Clear the variants
            VariantClear(&insertionPointVar);
        }
    }
    return output;
}

std::wstring handlePolyfaceMeshEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadPolyfaceMesh* mesh;  // Polyface Mesh interface pointer
    // Query for the IAcadPolyfaceMesh interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadPolyfaceMesh), (void**)&mesh);
    if (SUCCEEDED(hr))
    {
        if (mesh != nullptr) {
            // Initialize variants
            VARIANT coordinateVar;
            VariantInit(&coordinateVar);
            long numFacesVar, numVerticesVar;

            // Get details
            mesh->get_Coordinates(&coordinateVar);
            mesh->get_NumberOfFaces(&numFacesVar);
            mesh->get_NumberOfVertices(&numVerticesVar);

            // Assume that the returned variants are arrays of doubles and ints
            if (coordinateVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoordinate = coordinateVar.parray;
                double* coordinateArr;
                SafeArrayAccessData(saCoordinate, (void**)&coordinateArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Polyface Mesh has coordinates (";
                for (long i = 0; i < (saCoordinate->rgsabound[0].cElements); i++) {
                    if (i != 0)
                        std::wcout << L", ";
                    std::wcout << coordinateArr[i];
                }
                std::wcout << L")" << std::endl;

                if (numFacesVar == VT_I4 && numVerticesVar == VT_I4)
                    std::wcout << L"Number of Faces: " << numFacesVar << L", Number of Vertices: " << numVerticesVar << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoordinate);
            }

            // Clear the variants
            VariantClear(&coordinateVar);

        }
    }
    return output;
}

std::wstring handlePolygonMeshEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadPolygonMesh* polygonMesh;  // PolygonMesh interface pointer
    // Query for the IAcadPolygonMesh interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadPolygonMesh), (void**)&polygonMesh);

    if (SUCCEEDED(hr))
    {
        if (polygonMesh != nullptr) {
            // Initialize variants for properties
            VARIANT coordVar;
            VariantInit(&coordVar);


            // Get properties
            polygonMesh->get_Coordinates(&coordVar);

            // Assume that the returned variants are arrays of doubles
            if (coordVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoord = coordVar.parray;
                double* coordArr;
                SafeArrayAccessData(saCoord, (void**)&coordArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << " PolygonMesh with Coordinates (";
                for (long i = 0; i < saCoord->rgsabound[0].cElements; i++)
                    std::wcout << coordArr[i] << ((i != saCoord->rgsabound[0].cElements - 1) ? L", " : L")");

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);
                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoord);
            }

            // Clear the variants
            VariantClear(&coordVar);
        }
    }
    return output;
}

std::wstring handlePolylineEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadPolyline* polyline;  // Polyline interface pointer
    // Query for the IAcadPolyline interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadPolyline), (void**)&polyline);
    if (SUCCEEDED(hr))
    {
        if (polyline != nullptr)
        {
            double lengthVar, areaVar;
            VARIANT_BOOL closedVar;
            VARIANT coordVar;

            VariantInit(&coordVar);

            polyline->get_Area(&areaVar);
            polyline->get_Closed(&closedVar);
            polyline->get_Coordinates(&coordVar);
            polyline->get_Length(&lengthVar);

            // Assuming that the returned variants for coordinates is an array of doubles
            if (coordVar.vt == (VT_ARRAY | VT_R8))
            {
                SAFEARRAY* saCoord = coordVar.parray;
                double* coordArr;
                SafeArrayAccessData(saCoord, (void**)&coordArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Polyline with area: " << areaVar
                    << L", is closed: " << (closedVar ? L"Yes" : L"No")
                    << L", length: " << lengthVar
                    << L", coordinates: ";

                long lowerBound, upperBound;
                SafeArrayGetLBound(saCoord, 1, &lowerBound);
                SafeArrayGetUBound(saCoord, 1, &upperBound);

                for (long i = lowerBound; i <= upperBound; i += 2)
                {
                    std::wcout << L"(" << coordArr[i] << L", " << coordArr[i + 1] << L"), ";
                }

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoord);
            }

            // Clear the variants
            VariantClear(&coordVar);
        }
    }
    return output;
}

std::wstring handleNurbSurface(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    HRESULT hr;

    // Handle specific entity types
    if (entity.second != nullptr) {
        IAcadNurbSurface* nurbsSurface;  // NurbSurface interface pointer            
        // Query for the IAcadNurbSurface interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadNurbSurface), (void**)&nurbsSurface);

        if (SUCCEEDED(hr))
        {
            if (nurbsSurface != nullptr) {
                // Get some properties

                BSTR surfaceTypeVar;

                nurbsSurface->get_SurfaceType(&surfaceTypeVar);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" NURBS Surface with Surface Type " << surfaceTypeVar << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

            }
        }

        IAcadOle* oleObj;  // IAcadOle interface pointer
        // Query for the IAcadOle interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadOle), (void**)&oleObj);
        if (SUCCEEDED(hr))
        {
            if (oleObj != nullptr) {
                // Initialize variants for each property
                VARIANT oleInsertionPoint;
                VARIANT_BOOL  oleVisible;
                double oleHeight, oleWidth, oleScaleHeight, oleScaleWidth;
                BSTR oleEntityTransparency, oleOleSourceApp;
                AcOleType oleOleItemType;
                ACAD_ANGLE oleRotation;

                VariantInit(&oleInsertionPoint);

                // Get properties
                oleObj->get_Height(&oleHeight);
                oleObj->get_InsertionPoint(&oleInsertionPoint);
                oleObj->get_Width(&oleWidth);
                oleObj->get_EntityTransparency(&oleEntityTransparency);
                oleObj->get_OleItemType(&oleOleItemType);
                oleObj->get_OleSourceApp(&oleOleSourceApp);
                oleObj->get_Rotation(&oleRotation);
                oleObj->get_ScaleHeight(&oleScaleHeight);
                oleObj->get_ScaleWidth(&oleScaleWidth);
                oleObj->get_Visible(&oleVisible);

                // Print the properties
                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" OleObject with Height: " << oleHeight
                    << L", Width: " << oleWidth << L", Insertion Point: (" << oleInsertionPoint.dblVal << L"), Entity Transparency: "
                    << oleEntityTransparency << L", Ole Item Type: " << oleOleItemType << L", Ole Source App: "
                    << oleOleSourceApp << L", Rotation: " << oleRotation << L", Scale Height: " << oleScaleHeight
                    << L", Scale Width: " << oleScaleWidth << L", Visible: " << (oleVisible ? L"True" : L"False") << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Clear the variants
                VariantClear(&oleInsertionPoint);
            }
        }

        IAcadPlaneSurface* planeSurface;  // PlaneSurface interface pointer
        // Query for the IAcadPlaneSurface interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadPlaneSurface), (void**)&planeSurface);
        if (SUCCEEDED(hr))
        {
            if (planeSurface != nullptr) {
                // Initialize variants
                BSTR surfaceTypeVar, materialVar, lineTypeVar;

                // Get properties
                planeSurface->get_SurfaceType(&surfaceTypeVar);
                planeSurface->get_Material(&materialVar);
                planeSurface->get_Linetype(&lineTypeVar);


                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second
                    << L" PlaneSurface with SurfaceType " << surfaceTypeVar
                    << L", Material " << materialVar
                    << L", and LineType " << lineTypeVar << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;
            }
        }

        IAcadPoint* point;  // Point interface pointer
        // Query for the IAcadPoint interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadPoint), (void**)&point);
        if (SUCCEEDED(hr))
        {
            if (point != nullptr) {
                // Initialize variant
                VARIANT coordVar;
                VariantInit(&coordVar);

                // Get point coordinates
                point->get_Coordinates(&coordVar);

                // Assume that the returned variant is an array of doubles
                if (coordVar.vt == (VT_ARRAY | VT_R8)) {
                    SAFEARRAY* saCoords = coordVar.parray;

                    double* coordArr;

                    SafeArrayAccessData(saCoords, (void**)&coordArr);

                    std::wostringstream oss;
                    std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                    std::wcout << L"ID " << entity.first << ": " << entity.second << L" Point at (" << coordArr[0] << L", " << coordArr[1] << L", " << coordArr[2] << L")" << std::endl;

                    output = oss.str();
                    std::wcout.rdbuf(oldWcoutStreamBuf);

                    std::wcout << output << std::endl;

                    // Unaccess the data
                    SafeArrayUnaccessData(saCoords);
                }

                // Clear the variant
                VariantClear(&coordVar);
            }
        }

        IAcadPointCloud* pointCloud;  // PointCloud interface pointer
        // Query for the IAcadPointCloud interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadPointCloud), (void**)&pointCloud);
        if (SUCCEEDED(hr))
        {
            if (pointCloud != nullptr) {
                // Initialize variants
                BSTR nameVar;
                ACAD_DISTANCE heightVar, lengthVar, widthVar;
                VARIANT insertionPointVar;
                VariantInit(&insertionPointVar);

                // Get properties
                pointCloud->get_Name(&nameVar);
                pointCloud->get_Height(&heightVar);
                pointCloud->get_Length(&lengthVar);
                pointCloud->get_Width(&widthVar);
                pointCloud->get_InsertionPoint(&insertionPointVar);

                // Assume that the returned variants are arrays of doubles for insertion point
                if (insertionPointVar.vt == (VT_ARRAY | VT_R8)) {
                    SAFEARRAY* saInsertionPoint = insertionPointVar.parray;
                    double* insertionPointArr;
                    SafeArrayAccessData(saInsertionPoint, (void**)&insertionPointArr);

                    std::wostringstream oss;
                    std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                    std::wcout << L"ID " << entity.first << ": " << entity.second << L" PointCloud name: " << nameVar << L", Height: " << heightVar << L", Length: " << lengthVar
                        << L", Width: " << widthVar << L", InsertionPoint: (" << insertionPointArr[0] << L", " << insertionPointArr[1] << L", " << insertionPointArr[2] << L")" << std::endl;

                    output = oss.str();
                    std::wcout.rdbuf(oldWcoutStreamBuf);

                    std::wcout << output << std::endl;

                    // Unaccess the data
                    SafeArrayUnaccessData(saInsertionPoint);
                }

                // Clear the variants
                VariantClear(&insertionPointVar);
            }
        }

        IAcadPointCloudEx* pointCloudEx;  // PointCloud interface pointer
        // Query for the IAcadPointCloudEx interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadPointCloudEx), (void**)&pointCloudEx);

        if (SUCCEEDED(hr))
        {
            if (pointCloudEx != nullptr) {
                // Initialize variants
                BSTR name;
                VARIANT insertionPointVar;
                VariantInit(&insertionPointVar);

                // Get name and insertion point
                pointCloudEx->get_Name(&name);
                pointCloudEx->get_InsertionPoint(&insertionPointVar);

                // Assume that the returned variants are BSTR and arrays of doubles
                if (insertionPointVar.vt == (VT_ARRAY | VT_R8)) {
                    SAFEARRAY* saInsertion = insertionPointVar.parray;

                    double* insertionArr;

                    SafeArrayAccessData(saInsertion, (void**)&insertionArr);

                    std::wostringstream oss;
                    std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                    std::wcout << L"ID " << entity.first << ": " << entity.second << L" PointCloud '" << name << L"' with insertion point at (" << insertionArr[0] << L", " << insertionArr[1] << L", " << insertionArr[2] << L")" << std::endl;

                    output = oss.str();
                    std::wcout.rdbuf(oldWcoutStreamBuf);

                    std::wcout << output << std::endl;

                    // Unaccess the data
                    SafeArrayUnaccessData(saInsertion);
                }

                // Clear the variants
                VariantClear(&insertionPointVar);
            }
        }

        IAcadPolyfaceMesh* mesh;  // Polyface Mesh interface pointer
        // Query for the IAcadPolyfaceMesh interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadPolyfaceMesh), (void**)&mesh);
        if (SUCCEEDED(hr))
        {
            if (mesh != nullptr) {
                // Initialize variants
                VARIANT coordinateVar;
                VariantInit(&coordinateVar);
                long numFacesVar, numVerticesVar;

                // Get details
                mesh->get_Coordinates(&coordinateVar);
                mesh->get_NumberOfFaces(&numFacesVar);
                mesh->get_NumberOfVertices(&numVerticesVar);

                // Assume that the returned variants are arrays of doubles and ints
                if (coordinateVar.vt == (VT_ARRAY | VT_R8)) {
                    SAFEARRAY* saCoordinate = coordinateVar.parray;
                    double* coordinateArr;
                    SafeArrayAccessData(saCoordinate, (void**)&coordinateArr);

                    std::wostringstream oss;
                    std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                    std::wcout << L"ID " << entity.first << ": " << entity.second << L" Polyface Mesh has coordinates (";
                    for (long i = 0; i < (saCoordinate->rgsabound[0].cElements); i++) {
                        if (i != 0)
                            std::wcout << L", ";
                        std::wcout << coordinateArr[i];
                    }
                    std::wcout << L")" << std::endl;

                    if (numFacesVar == VT_I4 && numVerticesVar == VT_I4)
                        std::wcout << L"Number of Faces: " << numFacesVar << L", Number of Vertices: " << numVerticesVar << std::endl;

                    output = oss.str();
                    std::wcout.rdbuf(oldWcoutStreamBuf);

                    std::wcout << output << std::endl;

                    // Unaccess the data
                    SafeArrayUnaccessData(saCoordinate);
                }

                // Clear the variants
                VariantClear(&coordinateVar);

            }
        }

        IAcadPolygonMesh* polygonMesh;  // PolygonMesh interface pointer
        // Query for the IAcadPolygonMesh interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadPolygonMesh), (void**)&polygonMesh);

        if (SUCCEEDED(hr))
        {
            if (polygonMesh != nullptr) {
                // Initialize variants for properties
                VARIANT coordVar;
                VariantInit(&coordVar);


                // Get properties
                polygonMesh->get_Coordinates(&coordVar);

                // Assume that the returned variants are arrays of doubles
                if (coordVar.vt == (VT_ARRAY | VT_R8)) {
                    SAFEARRAY* saCoord = coordVar.parray;
                    double* coordArr;
                    SafeArrayAccessData(saCoord, (void**)&coordArr);

                    std::wostringstream oss;
                    std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                    std::wcout << L"ID " << entity.first << ": " << entity.second << " PolygonMesh with Coordinates (";
                    for (long i = 0; i < saCoord->rgsabound[0].cElements; i++)
                        std::wcout << coordArr[i] << ((i != saCoord->rgsabound[0].cElements - 1) ? L", " : L")");

                    output = oss.str();
                    std::wcout.rdbuf(oldWcoutStreamBuf);
                    std::wcout << output << std::endl;

                    // Unaccess the data
                    SafeArrayUnaccessData(saCoord);
                }

                // Clear the variants
                VariantClear(&coordVar);
            }
        }

        IAcadPolyline* polyline;  // Polyline interface pointer
        // Query for the IAcadPolyline interface from the entity
        hr = entity.second->QueryInterface(__uuidof(IAcadPolyline), (void**)&polyline);
        if (SUCCEEDED(hr))
        {
            if (polyline != nullptr)
            {
                double lengthVar, areaVar;
                VARIANT_BOOL closedVar;
                VARIANT coordVar;

                VariantInit(&coordVar);

                polyline->get_Area(&areaVar);
                polyline->get_Closed(&closedVar);
                polyline->get_Coordinates(&coordVar);
                polyline->get_Length(&lengthVar);

                // Assuming that the returned variants for coordinates is an array of doubles
                if (coordVar.vt == (VT_ARRAY | VT_R8))
                {
                    SAFEARRAY* saCoord = coordVar.parray;
                    double* coordArr;
                    SafeArrayAccessData(saCoord, (void**)&coordArr);

                    std::wostringstream oss;
                    std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                    std::wcout << L"ID " << entity.first << ": " << entity.second << L" Polyline with area: " << areaVar
                        << L", is closed: " << (closedVar ? L"Yes" : L"No")
                        << L", length: " << lengthVar
                        << L", coordinates: ";

                    long lowerBound, upperBound;
                    SafeArrayGetLBound(saCoord, 1, &lowerBound);
                    SafeArrayGetUBound(saCoord, 1, &upperBound);

                    for (long i = lowerBound; i <= upperBound; i += 2)
                    {
                        std::wcout << L"(" << coordArr[i] << L", " << coordArr[i + 1] << L"), ";
                    }

                    output = oss.str();
                    std::wcout.rdbuf(oldWcoutStreamBuf);

                    std::wcout << output << std::endl;

                    // Unaccess the data
                    SafeArrayUnaccessData(saCoord);
                }

                // Clear the variants
                VariantClear(&coordVar);
            }
        }
    }

    return output;
}

std::wstring handleViewportEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadPViewport* viewport;  // Viewport interface pointer
    // Query for the IAcadPViewport interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadPViewport), (void**)&viewport);
    if (SUCCEEDED(hr))
    {
        // Initialize variants
        VARIANT centerVar;
        VariantInit(&centerVar);
        double widthVar, heightVar;
        VARIANT directionVar;
        VariantInit(&directionVar);

        // Get properties
        viewport->get_Center(&centerVar);
        viewport->get_Height(&heightVar);
        viewport->get_Width(&widthVar);
        viewport->get_Direction(&directionVar);

        // Assume that the returned variants are arrays of doubles
        if (centerVar.vt == (VT_ARRAY | VT_R8) && directionVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saCenter = centerVar.parray;
            SAFEARRAY* saDirection = directionVar.parray;

            double* centerArr;
            double* directionArr;

            SafeArrayAccessData(saCenter, (void**)&centerArr);
            SafeArrayAccessData(saDirection, (void**)&directionArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second << L" Viewport with center (" << centerArr[0] << L", " << centerArr[1] << L", " << centerArr[2] << L") "
                << L"and direction (" << directionArr[0] << L", " << directionArr[1] << L", " << directionArr[2] << L") "
                << L"height: " << heightVar << L" width: " << widthVar << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saCenter);
            SafeArrayUnaccessData(saDirection);
        }

        // Clear the variants
        VariantClear(&centerVar);
        VariantClear(&directionVar);
    }
    return output;
}

std::wstring handleRasterImageEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadRasterImage* image;  // Image interface pointer
    // Query for the IAcadRasterImage interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadRasterImage), (void**)&image);

    if (SUCCEEDED(hr)) {
        // Initialize variants
        double widthVar, heightVar;
        VARIANT originVar;
        VariantInit(&originVar);
        BSTR imageFileVar;
        ACAD_ANGLE rotationVar;

        // Get properties
        image->get_Origin(&originVar);
        image->get_Height(&heightVar);
        image->get_Width(&widthVar);
        image->get_ImageFile(&imageFileVar);
        image->get_Rotation(&rotationVar);

        // Assume that the returned origin is an array of doubles
        if (originVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saOrigin = originVar.parray;
            double* originArr;
            SafeArrayAccessData(saOrigin, (void**)&originArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second
                << L" RasterImage origin (" << originArr[0] << L", " << originArr[1] << L", " << originArr[2]
                << L") height (" << heightVar << L") width (" << widthVar
                << L") imageFile (" << imageFileVar << L") rotation (" << rotationVar << L")" << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saOrigin);
        }

        // Clear the variants
        VariantClear(&originVar);
    }

    return output;
}

std::wstring handleRayEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadRay* ray;  // Ray interface pointer
    // Query for the IAcadRay interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadRay), (void**)&ray);
    if (SUCCEEDED(hr))
    {
        // Initialize variants
        VARIANT baseVar;
        VariantInit(&baseVar);
        VARIANT dirVar;
        VariantInit(&dirVar);

        // Get points
        ray->get_BasePoint(&baseVar);
        ray->get_DirectionVector(&dirVar);

        // Assume that the returned variants are arrays of doubles
        if (baseVar.vt == (VT_ARRAY | VT_R8) && dirVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saBase = baseVar.parray;
            SAFEARRAY* saDir = dirVar.parray;

            double* baseArr;
            double* dirArr;

            SafeArrayAccessData(saBase, (void**)&baseArr);
            SafeArrayAccessData(saDir, (void**)&dirArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second << L" Ray from (" << baseArr[0] << L", " << baseArr[1] << L", " << baseArr[2] << L") with direction vector ("
                << dirArr[0] << L", " << dirArr[1] << L", " << dirArr[2] << L")" << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saBase);
            SafeArrayUnaccessData(saDir);
        }

        // Clear the variants
        VariantClear(&baseVar);
        VariantClear(&dirVar);
    }
    return output;
}

std::wstring handleRegionEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadRegion* region;  // Region interface pointer
    // Query for the IAcadRegion interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadRegion), (void**)&region);
    if (SUCCEEDED(hr))
    {
        // Initialize variants
        double areaVar;
        VARIANT centroidVar;
        VariantInit(&centroidVar);

        // Get properties
        region->get_Area(&areaVar);
        region->get_Centroid(&centroidVar);

        // Assume that the returned variants are arrays of doubles
        if (centroidVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saCentroid = centroidVar.parray;

            double* centroidArr;

            SafeArrayAccessData(saCentroid, (void**)&centroidArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second << L" Region with Area: " << areaVar << " and Centroid at (" << centroidArr[0] << L", " << centroidArr[1] << L", " << centroidArr[2] << L")" << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saCentroid);
        }

        // Clear the variants
        VariantClear(&centroidVar);
    }
    return output;
}

std::wstring handleRevolvedSurfaceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadRevolvedSurface* revolvedSurface;  // RevolvedSurface interface pointer
    // Query for the IAcadRevolvedSurface interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadRevolvedSurface), (void**)&revolvedSurface);
    if (SUCCEEDED(hr)) {
        // Initialize variants
        VARIANT axisDirVar;
        VariantInit(&axisDirVar);
        VARIANT axisPosVar;
        VariantInit(&axisPosVar);
        ACAD_ANGLE revAngleVar;

        // Get properties
        revolvedSurface->get_AxisDirection(&axisDirVar);
        revolvedSurface->get_AxisPosition(&axisPosVar);
        revolvedSurface->get_RevolutionAngle(&revAngleVar);

        // Assume that the returned variants are arrays of doubles for coordinates
        if (axisDirVar.vt == (VT_ARRAY | VT_R8) && axisPosVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saAxisDir = axisDirVar.parray;
            SAFEARRAY* saAxisPos = axisPosVar.parray;

            double* axisDirArr;
            double* axisPosArr;

            SafeArrayAccessData(saAxisDir, (void**)&axisDirArr);
            SafeArrayAccessData(saAxisPos, (void**)&axisPosArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second
                << L" RevolvedSurface with AxisDirection (" << axisDirArr[0] << L", " << axisDirArr[1] << L", " << axisDirArr[2] << L") "
                << L"and AxisPosition (" << axisPosArr[0] << L", " << axisPosArr[1] << L", " << axisPosArr[2] << L") "
                << L"and RevolutionAngle " << revAngleVar << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saAxisDir);
            SafeArrayUnaccessData(saAxisPos);
        }

        // Clear the variants
        VariantClear(&axisDirVar);
        VariantClear(&axisPosVar);
    }
    return output;
}

std::wstring handleSectionEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadSection* section;  // Section interface pointer
    // Query for the IAcadSection interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadSection), (void**)&section);
    if (SUCCEEDED(hr))
    {
        // Initialize variants
        VARIANT coordinateVar;
        VariantInit(&coordinateVar);
        BSTR nameVar;
        double elevationVar;

        // Get properties
        section->get_Coordinate(0, &coordinateVar);
        section->get_Name(&nameVar);
        section->get_Elevation(&elevationVar);

        // Assume that the returned coordinate variant is an array of doubles
        if (coordinateVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saCoordinate = coordinateVar.parray;

            double* coordinateArr;

            SafeArrayAccessData(saCoordinate, (void**)&coordinateArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second << L" Section with name: " << nameVar
                << L", elevation: " << elevationVar
                << L", and coordinate (" << coordinateArr[0] << L", " << coordinateArr[1] << L", " << coordinateArr[2] << L")" << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saCoordinate);
        }

        // Clear the variants
        VariantClear(&coordinateVar);
    }
    return output;
}

std::wstring handleShapeEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadShape* shape;  // Shape interface pointer
    // Query for the IAcadShape interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadShape), (void**)&shape);
    if (SUCCEEDED(hr))
    {
        BSTR name;
        VARIANT insPointVar;
        VariantInit(&insPointVar);
        double height;
        double rotation;
        double scaleFactor;

        // Get shape properties
        shape->get_Name(&name);
        shape->get_InsertionPoint(&insPointVar);
        shape->get_Height(&height);
        shape->get_Rotation(&rotation);
        shape->get_ScaleFactor(&scaleFactor);

        // Assume that the returned variant is an array of doubles
        if (insPointVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saInsPoint = insPointVar.parray;

            double* insPointArr;

            SafeArrayAccessData(saInsPoint, (void**)&insPointArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second << L" Shape name: " << name << L", Insertion point: (" << insPointArr[0] << L", " << insPointArr[1] << L", " << insPointArr[2] << L"), Height: "
                << height << L", Rotation: " << rotation << L", Scale factor: " << scaleFactor << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saInsPoint);
        }

        // Clear the variant
        VariantClear(&insPointVar);
        // Release the BSTR
        SysFreeString(name);
    }
    return output;
}

std::wstring handleSolidEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadSolid* solid;  // Solid interface pointer
    // Query for the IAcadSolid interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadSolid), (void**)&solid);
    if (SUCCEEDED(hr))
    {
        // Initialize variants
        VARIANT coordVar;
        VariantInit(&coordVar);
        VARIANT normalVar;
        VariantInit(&normalVar);

        // Get points
        solid->get_Coordinates(&coordVar);
        solid->get_Normal(&normalVar);

        // Assume that the returned variants are arrays of doubles
        if (coordVar.vt == (VT_ARRAY | VT_R8) && normalVar.vt == (VT_ARRAY | VT_R8)) {
            SAFEARRAY* saCoord = coordVar.parray;
            SAFEARRAY* saNormal = normalVar.parray;

            double* coordArr;
            double* normalArr;

            SafeArrayAccessData(saCoord, (void**)&coordArr);
            SafeArrayAccessData(saNormal, (void**)&normalArr);

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second << L" Solid with Coordinates ("
                << coordArr[0] << L", " << coordArr[1] << L", " << coordArr[2] << L") and Normal ("
                << normalArr[0] << L", " << normalArr[1] << L", " << normalArr[2] << L")" << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Unaccess the data
            SafeArrayUnaccessData(saCoord);
            SafeArrayUnaccessData(saNormal);
        }

        // Clear the variants
        VariantClear(&coordVar);
        VariantClear(&normalVar);
    }
    return output;
}

std::wstring handleSplineEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadSpline* spline;  // Spline interface pointer
    // Query for the IAcadSpline interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadSpline), (void**)&spline);
    if (SUCCEEDED(hr))
    {
        if (spline != nullptr)
        {
            // Initialize variants for ControlPoints and FitPoints
            VARIANT ctrlPointsVar;
            VariantInit(&ctrlPointsVar);
            VARIANT fitPointsVar;
            VariantInit(&fitPointsVar);

            // Get points
            spline->get_ControlPoints(&ctrlPointsVar);
            spline->get_FitPoints(&fitPointsVar);

            // Assume that the returned variants are arrays of doubles
            if (ctrlPointsVar.vt == (VT_ARRAY | VT_R8) && fitPointsVar.vt == (VT_ARRAY | VT_R8))
            {
                SAFEARRAY* saCtrlPoints = ctrlPointsVar.parray;
                SAFEARRAY* saFitPoints = fitPointsVar.parray;

                double* ctrlPointsArr;
                double* fitPointsArr;

                SafeArrayAccessData(saCtrlPoints, (void**)&ctrlPointsArr);
                SafeArrayAccessData(saFitPoints, (void**)&fitPointsArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                // Output control points and fit points. Assuming 3D points, stride of 3.
                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Spline control points: ";
                for (long i = 0; i < (saCtrlPoints->rgsabound[0].cElements / 3); i++)
                {
                    std::wcout << L"(" << ctrlPointsArr[i * 3] << L", " << ctrlPointsArr[i * 3 + 1] << L", " << ctrlPointsArr[i * 3 + 2] << L"), ";
                }

                std::wcout << L" Fit points: ";
                for (long i = 0; i < (saFitPoints->rgsabound[0].cElements / 3); i++)
                {
                    std::wcout << L"(" << fitPointsArr[i * 3] << L", " << fitPointsArr[i * 3 + 1] << L", " << fitPointsArr[i * 3 + 2] << L"), ";
                }
                std::wcout << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCtrlPoints);
                SafeArrayUnaccessData(saFitPoints);
            }

            // Clear the variants
            VariantClear(&ctrlPointsVar);
            VariantClear(&fitPointsVar);
        }
    }
    return output;
}

std::wstring handleSubDMeshEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadSubDMesh* subDMesh;  // SubDMesh interface pointer
    // Query for the IAcadSubDMesh interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadSubDMesh), (void**)&subDMesh);
    if (SUCCEEDED(hr))
    {
        if (subDMesh != nullptr) {
            // Initialize variant
            VARIANT coordVar;
            VariantInit(&coordVar);

            // Get coordinates
            subDMesh->get_Coordinates(&coordVar);

            // Assume that the returned variant is an array of doubles
            if (coordVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoord = coordVar.parray;

                double* coordArr;

                SafeArrayAccessData(saCoord, (void**)&coordArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Mesh with coordinates: ";

                // Assume 3D points
                for (long i = 0; i < saCoord->rgsabound[0].cElements; i += 3)
                {
                    std::wcout << L"(" << coordArr[i] << L", " << coordArr[i + 1] << L", " << coordArr[i + 2] << L"), ";
                }

                std::wcout << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoord);
            }

            // Clear the variant
            VariantClear(&coordVar);
        }
    }
    return output;
}

std::wstring handleSurfaceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadSurface* surface;  // Surface interface pointer
    // Query for the IAcadSurface interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadSurface), (void**)&surface);
    if (SUCCEEDED(hr))
    {
        if (surface != nullptr) {
            // Get properties
            BSTR layer;
            LPDISPATCH application, document;

            surface->get_Application(&application);
            surface->get_Document(&document);
            surface->get_Layer(&layer);


            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second
                << L"Surface with Application: " << (wchar_t*)application
                << L" Document: " << (wchar_t*)document
                << L" Layer: " << (wchar_t*)layer << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;

            // Release BSTRs
            SysFreeString(layer);
        }
    }
    return output;
}

std::wstring handleSweptSurfaceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadSweptSurface* sweptSurface;  // Swept Surface interface pointer
    // Query for the IAcadSweptSurface interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadSweptSurface), (void**)&sweptSurface);
    if (SUCCEEDED(hr))
    {
        if (sweptSurface != nullptr) {
            // Initialize variants
            double lenVar, scaleVar;

            // Get properties
            sweptSurface->get_Scale(&scaleVar);
            sweptSurface->get_Length(&lenVar);

            // Check that the returned variants are of the correct type
            double scale = scaleVar;
            double length = lenVar;

            std::wostringstream oss;
            std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

            std::wcout << L"ID " << entity.first << ": " << entity.second << L" SweptSurface with Scale: " << scale << L", Length: " << length << std::endl;

            output = oss.str();
            std::wcout.rdbuf(oldWcoutStreamBuf);

            std::wcout << output << std::endl;
        }
    }
    return output;
}

std::wstring handleTableEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadTable* table;  // Table interface pointer
    // Query for the IAcadTable interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadTable), (void**)&table);
    if (SUCCEEDED(hr))
    {
        if (table != nullptr) {
            // Initialize variants
            VARIANT insertionPointVar;
            VariantInit(&insertionPointVar);
            double heightVar, widthVar;

            // Get properties
            table->get_InsertionPoint(&insertionPointVar);
            table->get_Height(&heightVar);
            table->get_Width(&widthVar);

            // Assume that the returned variants are arrays of doubles for InsertionPoint
            if (insertionPointVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saInsertionPoint = insertionPointVar.parray;
                double* insertPointArr;
                SafeArrayAccessData(saInsertionPoint, (void**)&insertPointArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Table at (" << insertPointArr[0] << L", " << insertPointArr[1] << L", " << insertPointArr[2] << L") "
                    << L"Width: " << widthVar << L", Height: " << heightVar << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saInsertionPoint);
            }

            // Clear the variants
            VariantClear(&insertionPointVar);
        }
    }
    return output;
}

std::wstring handleTextEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadText* text;  // Text interface pointer
    // Query for the IAcadText interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadText), (void**)&text);
    if (SUCCEEDED(hr))
    {
        if (text != nullptr) {
            // Initialize variants
            VARIANT insertionPointVar;
            VariantInit(&insertionPointVar);
            double height;
            double rotation;
            BSTR textString;

            // Get properties
            text->get_InsertionPoint(&insertionPointVar);
            text->get_Height(&height);
            text->get_Rotation(&rotation);
            text->get_TextString(&textString);

            // Assume that the returned variants are arrays of doubles
            if (insertionPointVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saInsert = insertionPointVar.parray;
                double* insertArr;
                SafeArrayAccessData(saInsert, (void**)&insertArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Text inserted at (" << insertArr[0] << L", " << insertArr[1] << L", " << insertArr[2] << L"), "
                    << L"Height: " << height << L", Rotation: " << rotation << L", Text: " << textString << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saInsert);
            }

            // Clear the variants
            VariantClear(&insertionPointVar);
        }
    }
    return output;
}

std::wstring handleToleranceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadTolerance* tolerance;  // Tolerance interface pointer
    // Query for the IAcadTolerance interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadTolerance), (void**)&tolerance);
    if (SUCCEEDED(hr))
    {
        if (tolerance != nullptr) {
            // Initialize variants
            VARIANT insPointVar;
            VariantInit(&insPointVar);
            VARIANT dirVectorVar;
            VariantInit(&dirVectorVar);

            // Get properties
            tolerance->get_InsertionPoint(&insPointVar);
            tolerance->get_DirectionVector(&dirVectorVar);

            // Assume that the returned variants are arrays of doubles
            if (insPointVar.vt == (VT_ARRAY | VT_R8) && dirVectorVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saInsPoint = insPointVar.parray;
                SAFEARRAY* saDirVector = dirVectorVar.parray;

                double* insPointArr;
                double* dirVectorArr;

                SafeArrayAccessData(saInsPoint, (void**)&insPointArr);
                SafeArrayAccessData(saDirVector, (void**)&dirVectorArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Tolerance with Insertion Point (" << insPointArr[0] << L", " << insPointArr[1] << L", " << insPointArr[2]
                    << L") and Direction Vector (" << dirVectorArr[0] << L", " << dirVectorArr[1] << L", " << dirVectorArr[2] << L")" << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saInsPoint);
                SafeArrayUnaccessData(saDirVector);
            }

            // Clear the variants
            VariantClear(&insPointVar);
            VariantClear(&dirVectorVar);
        }
    }
    return output;
}

std::wstring handleTraceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadTrace* trace;  // Trace interface pointer
    // Query for the IAcadTrace interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadTrace), (void**)&trace);
    if (SUCCEEDED(hr))
    {
        if (trace != nullptr) {
            // Initialize variants
            VARIANT coordsVar;
            VariantInit(&coordsVar);
            VARIANT normalVar;
            VariantInit(&normalVar);

            // Get points and normal
            trace->get_Coordinates(&coordsVar);
            trace->get_Normal(&normalVar);

            // Assume that the returned variants are arrays of doubles
            if (coordsVar.vt == (VT_ARRAY | VT_R8) && normalVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoords = coordsVar.parray;
                SAFEARRAY* saNormal = normalVar.parray;

                double* coordsArr;
                double* normalArr;

                SafeArrayAccessData(saCoords, (void**)&coordsArr);
                SafeArrayAccessData(saNormal, (void**)&normalArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Trace with Coordinates: (";
                for (long i = 0; i < saCoords->rgsabound[0].cElements; i++) {
                    std::wcout << coordsArr[i];
                    if (i < saCoords->rgsabound[0].cElements - 1) {
                        std::wcout << L", ";
                    }
                }
                std::wcout << L") and Normal: (" << normalArr[0] << L", " << normalArr[1] << L", " << normalArr[2] << L")" << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoords);
                SafeArrayUnaccessData(saNormal);
            }

            // Clear the variants
            VariantClear(&coordsVar);
            VariantClear(&normalVar);
        }
    }
    return output;
}

std::wstring handleWipeoutEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadWipeout* wipeout;  // Wipeout interface pointer
    // Query for the IAcadWipeout interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadWipeout), (void**)&wipeout);
    if (SUCCEEDED(hr))
    {
        if (wipeout != nullptr) {
            // Initialize variants
            VARIANT originVar;
            VariantInit(&originVar);
            double height, width;
            BSTR imageFile;

            // Get properties
            wipeout->get_Origin(&originVar);
            wipeout->get_Height(&height);
            wipeout->get_Width(&width);
            wipeout->get_ImageFile(&imageFile);

            // Assume that the returned variant is an array of doubles
            if (originVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saOrigin = originVar.parray;

                double* originArr;

                SafeArrayAccessData(saOrigin, (void**)&originArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Wipeout from (" << originArr[0] << L", " << originArr[1] << L", " << originArr[2] << L") with dimensions "
                    << L"Height: " << height << L", Width: " << width << L" and ImageFile: " << imageFile << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saOrigin);
            }

            // Clear the variant
            VariantClear(&originVar);
        }
    }
    return output;
}

std::wstring handleXlineEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcadXline* xline;  // Xline interface pointer
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcadXline), (void**)&xline);
    if (SUCCEEDED(hr))
    {
        if (xline != nullptr) {
            // Initialize variants
            VARIANT basePointVar;
            VariantInit(&basePointVar);
            VARIANT directionVar;
            VariantInit(&directionVar);
            VARIANT secondPointVar;
            VariantInit(&secondPointVar);

            // Get properties
            xline->get_BasePoint(&basePointVar);
            xline->get_DirectionVector(&directionVar);
            xline->get_SecondPoint(&secondPointVar);

            // Assume that the returned variants are arrays of doubles
            if (basePointVar.vt == (VT_ARRAY | VT_R8) && directionVar.vt == (VT_ARRAY | VT_R8) && secondPointVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saBasePoint = basePointVar.parray;
                SAFEARRAY* saDirection = directionVar.parray;
                SAFEARRAY* saSecondPoint = secondPointVar.parray;

                double* basePointArr;
                double* directionArr;
                double* secondPointArr;

                SafeArrayAccessData(saBasePoint, (void**)&basePointArr);
                SafeArrayAccessData(saDirection, (void**)&directionArr);
                SafeArrayAccessData(saSecondPoint, (void**)&secondPointArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << L": " << entity.second << L" Xline with base point (" << basePointArr[0] << L", " << basePointArr[1] << L", " << basePointArr[2] << L"), direction vector ("
                    << directionArr[0] << L", " << directionArr[1] << L", " << directionArr[2] << L"), and second point ("
                    << secondPointArr[0] << L", " << secondPointArr[1] << L", " << secondPointArr[2] << L")" << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saBasePoint);
                SafeArrayUnaccessData(saDirection);
                SafeArrayUnaccessData(saSecondPoint);
            }

            // Clear the variants
            VariantClear(&basePointVar);
            VariantClear(&directionVar);
            VariantClear(&secondPointVar);
        }
    }
    return output;
}

std::wstring handle3DSolidEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcad3DSolid* Dsolid;  // Solid interface pointer
    // Query for the IAcad3DSolid interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcad3DSolid), (void**)&Dsolid);
    if (SUCCEEDED(hr))
    {
        if (Dsolid != nullptr) {
            // Initialize variants
            VARIANT positionVar;
            VariantInit(&positionVar);
            VARIANT centroidVar;
            VariantInit(&centroidVar);
            double volumeVar;

            // Get properties
            Dsolid->get_Position(&positionVar);
            Dsolid->get_Centroid(&centroidVar);
            Dsolid->get_Volume(&volumeVar);

            // Assume that the returned variants are arrays of doubles for position and centroid, and a double for volume
            if (positionVar.vt == (VT_ARRAY | VT_R8) && centroidVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saPosition = positionVar.parray;
                SAFEARRAY* saCentroid = centroidVar.parray;
                double volume = volumeVar;

                double* positionArr;
                double* centroidArr;

                SafeArrayAccessData(saPosition, (void**)&positionArr);
                SafeArrayAccessData(saCentroid, (void**)&centroidArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" Solid at position (" << positionArr[0] << L", " << positionArr[1] << L", " << positionArr[2] << L"), "
                    << L"centroid at (" << centroidArr[0] << L", " << centroidArr[1] << L", " << centroidArr[2] << L"), "
                    << L"with volume " << volume << std::endl;

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saPosition);
                SafeArrayUnaccessData(saCentroid);
            }

            // Clear the variants
            VariantClear(&positionVar);
            VariantClear(&centroidVar);
        }
    }
    return output;
}

std::wstring handle3DPolylineEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcad3DPolyline* Dpolyline;  // Polyline interface pointer
    // Query for the IAcad3DPolyline interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcad3DPolyline), (void**)&Dpolyline);
    if (SUCCEEDED(hr))
    {
        if (Dpolyline != nullptr) {
            // Initialize variant
            VARIANT coordVar;
            VariantInit(&coordVar);
            VARIANT_BOOL closedStatus;

            // Get points and closed status
            Dpolyline->get_Coordinates(&coordVar);
            Dpolyline->get_Closed(&closedStatus);

            // Assume that the returned variant is an array of doubles
            if (coordVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoord = coordVar.parray;

                double* coordArr;

                SafeArrayAccessData(saCoord, (void**)&coordArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                // Output polyline ID and whether it's closed
                std::wcout << L"ID " << entity.first << ": " << entity.second << L" 3D Polyline, Closed: " << (closedStatus ? L"Yes" : L"No") << std::endl;

                // Output the coordinates
                long lower, upper;
                SafeArrayGetLBound(saCoord, 1, &lower);
                SafeArrayGetUBound(saCoord, 1, &upper);
                for (long i = lower; i <= upper; i += 3)
                {
                    std::wcout << L"Point: (" << coordArr[i] << L", " << coordArr[i + 1] << L", " << coordArr[i + 2] << L")" << std::endl;
                }

                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoord);
            }

            // Clear the variant
            VariantClear(&coordVar);
        }
    }
    return output;
}

std::wstring handle3DFaceEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity) {
    std::wstring output;
    IAcad3DFace* Dface;  // 3DFace interface pointer
    // Query for the IAcad3DFace interface from the entity
    HRESULT hr = entity.second->QueryInterface(__uuidof(IAcad3DFace), (void**)&Dface);
    if (SUCCEEDED(hr))
    {
        if (Dface != nullptr) {
            // Initialize variants
            VARIANT coordinatesVar;
            VariantInit(&coordinatesVar);

            // Get points
            Dface->get_Coordinates(&coordinatesVar);

            // Assume that the returned variant is an array of doubles
            if (coordinatesVar.vt == (VT_ARRAY | VT_R8)) {
                SAFEARRAY* saCoordinates = coordinatesVar.parray;
                double* coordinatesArr;
                SafeArrayAccessData(saCoordinates, (void**)&coordinatesArr);

                std::wostringstream oss;
                std::wstreambuf* oldWcoutStreamBuf = std::wcout.rdbuf(oss.rdbuf());

                std::wcout << L"ID " << entity.first << ": " << entity.second << L" 3DFace coordinates: (";

                // Assuming the array of doubles will contain an even number of elements
                long lowerBound, upperBound;
                SafeArrayGetLBound(saCoordinates, 1, &lowerBound);
                SafeArrayGetUBound(saCoordinates, 1, &upperBound);

                for (long i = lowerBound; i <= upperBound; i += 3) {
                    std::wcout << L"(" << coordinatesArr[i] << L", " << coordinatesArr[i + 1] << L", " << coordinatesArr[i + 2] << L")";
                    if (i < upperBound - 2) {
                        std::wcout << L", ";
                    }
                }

                std::wcout << L")" << std::endl;
                output = oss.str();
                std::wcout.rdbuf(oldWcoutStreamBuf);

                std::wcout << output << std::endl;

                // Unaccess the data
                SafeArrayUnaccessData(saCoordinates);
            }

            // Clear the variant
            VariantClear(&coordinatesVar);
        }
    }
    return output;
}

std::wstring handleEntity(const std::pair<std::wstring, CComPtr<IAcadEntity>>& entity, std::wstring formattedEntityName) {
    std::wstring output;

    if (formattedEntityName == L"line")
        output = handleLineEntity(entity);
    else if (formattedEntityName == L"arc")
        output = handleArcEntity(entity);
    else if (formattedEntityName == L"attribute")
        output = handleAttributeEntity(entity);
    else if (formattedEntityName == L"viewrepblockreference")
        output = handleBlockReferenceEntity(entity);
    else if (formattedEntityName == L"blockreference")
        output = handleBlock(entity);
    else if (formattedEntityName == L"circle")
        output = handleCircleEntity(entity);
    else if (formattedEntityName == L"ellipse")
        output = handleEllipseEntity(entity);
    else if (formattedEntityName == L"extruded_surface")
        output = handleExtrudedSurfaceEntity(entity);
    else if (formattedEntityName == L"hatch")
        output = handleHatchEntity(entity);
    else if (formattedEntityName == L"helix")
        output = handleHelixEntity(entity);
    else if (formattedEntityName == L"leader")
        output = handleLeaderEntity(entity);
    else if (formattedEntityName == L"lofted_surface")
        output = handleLoftedSurfaceEntity(entity);
    else if (formattedEntityName == L"minsert_block")
        output = handleMInsertBlockEntity(entity);
    else if (formattedEntityName == L"mleader")
        output = handleMLeaderEntity(entity);
    else if (formattedEntityName == L"mline")
        output = handleMLineEntity(entity);
    else if (formattedEntityName == L"mtext")
        output = handleMTextEntity(entity);
    else if (formattedEntityName == L"nurbs_surface")
        output = handleNurbsSurfaceEntity(entity);
    else if (formattedEntityName == L"ole_object")
        output = handleOleObjectEntity(entity);
    else if (formattedEntityName == L"plane_surface")
        output = handlePlaneSurfaceEntity(entity);
    else if (formattedEntityName == L"point")
        output = handlePointEntity(entity);
    else if (formattedEntityName == L"point_cloud")
        output = handlePointCloudEntity(entity);
    else if (formattedEntityName == L"point_cloud_ex")
        output = handlePointCloudExEntity(entity);
    else if (formattedEntityName == L"polyface_mesh")
        output = handlePolyfaceMeshEntity(entity);
    else if (formattedEntityName == L"polygon_mesh")
        output = handlePolygonMeshEntity(entity);
    else if (formattedEntityName == L"polyline")
        output = handlePolylineEntity(entity);
    else if (formattedEntityName == L"nurbsurface")
        output = handleNurbSurface(entity);
    else if (formattedEntityName == L"viewport")
        output = handleViewportEntity(entity);
    else if (formattedEntityName == L"raster_image")
        output = handleRasterImageEntity(entity);
    else if (formattedEntityName == L"ray")
        output = handleRayEntity(entity);
    else if (formattedEntityName == L"region")
        output = handleRegionEntity(entity);
    else if (formattedEntityName == L"revolved_surface")
        output = handleRevolvedSurfaceEntity(entity);
    else if (formattedEntityName == L"section")
        output = handleSectionEntity(entity);
    else if (formattedEntityName == L"shape")
        output = handleShapeEntity(entity);
    else if (formattedEntityName == L"solid")
        output = handleSolidEntity(entity);
    else if (formattedEntityName == L"spline")
        output = handleSplineEntity(entity);
    else if (formattedEntityName == L"subd_mesh")
        output = handleSubDMeshEntity(entity);
    else if (formattedEntityName == L"surface")
        output = handleSurfaceEntity(entity);
    else if (formattedEntityName == L"swept_surface")
        output = handleSweptSurfaceEntity(entity);
    else if (formattedEntityName == L"table")
        output = handleTableEntity(entity);
    else if (formattedEntityName == L"text")
        output = handleTextEntity(entity);
    else if (formattedEntityName == L"tolerance")
        output = handleToleranceEntity(entity);
    else if (formattedEntityName == L"trace")
        output = handleTraceEntity(entity);
    else if (formattedEntityName == L"wipeout")
        output = handleWipeoutEntity(entity);
    else if (formattedEntityName == L"xline")
        output = handleXlineEntity(entity);
    else if (formattedEntityName == L"3dsolid")
        output = handle3DSolidEntity(entity);
    else if (formattedEntityName == L"3dpolyline")
        output = handle3DPolylineEntity(entity);
    else if (formattedEntityName == L"3dface")
        output = handle3DFaceEntity(entity);
    return output;
}

void InspectBlockEntities(IAcadBlockReference* pAcadBlockRef, std::map<std::wstring, IAcadEntity*>& entitiesMap, std::set<std::wstring>& entitiesNames) {
    BSTR blockName;
    pAcadBlockRef->get_Name(&blockName);
    //std::wcout << "Block name: " << (wchar_t*)blockName << std::endl;

    // Release the block reference
    pAcadBlockRef->Release();

    // Store the block name in the entitiesNames set
    std::wstring formattedBlockName = getFormattedEntityName(blockName);
    entitiesNames.insert(formattedBlockName);
}

void InspectEntities(IAcadModelSpace* pAcadModelSpace, std::map<std::wstring, IAcadEntity*>& entitiesMap, std::set<std::wstring>& entitiesNames) {
    long numEntities;
    pAcadModelSpace->get_Count(&numEntities);

    for (long i = 0; i < numEntities; i++) {
        IAcadEntity* pAcadEntity;
        VARIANT index;
        index.vt = VT_I4;
        index.lVal = i;
        HRESULT hr = pAcadModelSpace->Item(index, &pAcadEntity);
        if (SUCCEEDED(hr)) {
            BSTR entityHandle;
            pAcadEntity->get_Handle(&entityHandle);

            long entityType;
            pAcadEntity->get_EntityType(&entityType);

            BSTR entityName;
            pAcadEntity->get_EntityName(&entityName);
            std::wstring formattedEntityName = getFormattedEntityName(entityName);
            entitiesNames.insert(formattedEntityName);

            if (entityType == acBlockReference) {
                IAcadBlockReference* pAcadBlockRef;
                pAcadEntity->QueryInterface(__uuidof(IAcadBlockReference), (void**)&pAcadBlockRef);
                if (pAcadBlockRef) {
                    InspectBlockEntities(pAcadBlockRef, entitiesMap, entitiesNames);
                }
            }

            // Store the entity in the map
            std::wstring key = std::wstring((wchar_t*)entityHandle);
            entitiesMap[key] = pAcadEntity;
        }
    }
}

Acad::ErrorStatus removeEntity(AcDbObjectId objectId)
{
    Acad::ErrorStatus es;

    AcDbEntity* pEntity;
    es = acdbOpenObject(pEntity, objectId, AcDb::kForWrite);

    if (es != Acad::eOk)
        return es;

    es = pEntity->erase(Adesk::kTrue); // Erase the entity
    pEntity->close();

    return es;
}


int _tmain(int argc, _TCHAR* argv[]) {
    // Локализация 
    setlocale(LC_ALL, "Russian");
    std::wcout.imbue(std::locale("Russian"));
    SetConsoleOutputCP(CP_UTF8);

    std::wcout << L"Введите путь к .dwg файлу: ";
    std::wstring pathInput;
    std::getline(std::wcin, pathInput);

    std::wcout << L"Введите путь, по которому будет сохранен .txt файл: ";
    std::wstring outputFile;
    std::getline(std::wcin, outputFile);

    outputFile += L"output.txt";
//    std::wcout << pathInput.c_str();
    // Initialize COM
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    if (SUCCEEDED(hr)) {
        DebugPrint(_T("COM initialization successful"));

        // Get a pointer to AutoCAD application
        IAcadApplication* pAcadApp;
        hr = CoCreateInstance(
            __uuidof(AcadApplication), NULL,
            CLSCTX_LOCAL_SERVER, __uuidof(IAcadApplication),
            (void**)&pAcadApp);

        if (SUCCEEDED(hr)) {
            DebugPrint(_T("AutoCAD application pointer obtained"));

            // Start AutoCAD
            pAcadApp->put_Visible(VARIANT_FALSE); // AutoCAD will be launched in the background
            DebugPrint(_T("AutoCAD started"));

            // Wait until AutoCAD is idle
            long long hwnd;
            pAcadApp->get_HWND(&hwnd);
            DWORD processId;
            GetWindowThreadProcessId((HWND)hwnd, &processId);
            HANDLE hProcess = OpenProcess(SYNCHRONIZE, FALSE, processId);
            WaitForInputIdle(hProcess, INFINITE);
            CloseHandle(hProcess);

            // Wait until AutoCAD is ready
            IAcadState* pAcadState;
            pAcadApp->GetAcadState(&pAcadState);
            VARIANT_BOOL isQuiescent;
            int timeout = 30;  // Timeout after 30 seconds
            do {
                Sleep(1000);
                pAcadState->get_IsQuiescent(&isQuiescent);
            } while (isQuiescent == VARIANT_FALSE && --timeout > 0);

            if (timeout <= 0) {
                DebugPrint(_T("Timeout waiting for AutoCAD to become ready"));
                // Handle the timeout situation as you see fit
            }


            // Get the Documents collection
            IAcadDocuments* pAcadDocs;
            hr = pAcadApp->get_Documents(&pAcadDocs);

            if (SUCCEEDED(hr)) {
                DebugPrint(_T("Documents collection obtained"));

                // Load DWG file
                BSTR bstrPath = SysAllocString(pathInput.c_str()); // Path to your DWG file
                IAcadDocument* pAcadDoc;
                VARIANT vReadOnly;
                vReadOnly.vt = VT_BOOL;
                vReadOnly.boolVal = VARIANT_TRUE;
                VARIANT vPassword;
                vPassword.vt = VT_EMPTY;
                hr = pAcadDocs->Open(bstrPath, vReadOnly, vPassword, &pAcadDoc);

                if (SUCCEEDED(hr)) {
                    DebugPrint(_T("DWG file loaded"));
                    Sleep(5000);

                    // Perform some operations...
                    IAcadModelSpace* pAcadModelSpace;
                    hr = pAcadDoc->get_ModelSpace(&pAcadModelSpace);
                    if (SUCCEEDED(hr)) {
                        DebugPrint(_T("ModelSpace obtained"));

                        long numEntities;
                        pAcadModelSpace->get_Count(&numEntities);

                        DebugPrint(_T("Iterating through entities"));
                        //std::map<std::wstring, IAcadEntity*> entitiesMap;
                        //std::set<std::wstring> entitiesNames;
                        for (long i = 0; i < numEntities; i++) {
                            IAcadEntity* pAcadEntity;
                            VARIANT index;
                            index.vt = VT_I4;
                            index.lVal = i;
                            hr = pAcadModelSpace->Item(index, &pAcadEntity);
                            if (SUCCEEDED(hr)) {
                                BSTR entityHandle;
                                pAcadEntity->get_Handle(&entityHandle);
                                //std::wcout << "Entity handle: " << (wchar_t*)entityHandle << std::endl;

                                long entityType;
                                pAcadEntity->get_EntityType(&entityType);
                                //std::wcout << "Entity type: " << entityType << std::endl;

                                BSTR entityName;
                                pAcadEntity->get_EntityName(&entityName);
                                std::wstring formattedEntityName = getFormattedEntityName(entityName);
                                entitiesNames.insert(formattedEntityName);
                                //std::wcout << "Entity name: " << formattedEntityName << std::endl;


                                if (entityType == acBlockReference) { // Check if it's a block reference
                                    IAcadBlockReference* pAcadBlockRef;
                                    pAcadEntity->QueryInterface(__uuidof(IAcadBlockReference), (void**)&pAcadBlockRef);
                                    if (pAcadBlockRef) {

                                        BSTR blockName;
                                        pAcadBlockRef->get_Name(&blockName);
                                        //std::wcout << "Block name: " << (wchar_t*)blockName << std::endl;
                                        InspectEntities(pAcadModelSpace, entitiesMap, entitiesNames);
                                        // Release the block reference
                                        pAcadBlockRef->Release();
                                    }
                                }


                                // Store the entity in the map
                                std::wstring key = std::wstring((wchar_t*)entityHandle);
                                entitiesMap[key] = pAcadEntity;

                                // No need to release pAcadEntity here, it will be released later after user's input
                            }
                        }

                        // Ask the user to input the handle of the entity they want to work with
                        while (true) {
                            std::wcout << L"Выберите блок или примитив, с которым желаете работать: \n";
                            for (const auto& name : entitiesNames) {
                                std::wcout << name << ", ";
                            }
                            std::cout << std::endl;
                            std::wstring userInputName;
                            std::getline(std::wcin, userInputName);

                            // Ищем объекты, с которыми надо взаимодействовать
                            std::map<std::wstring, IAcadEntity*> workEntitiesMap;
                            for (auto &entity : entitiesMap) {
                                IAcadEntity* pAcadEntity = entity.second;
                                BSTR entityName;
                                pAcadEntity->get_EntityName(&entityName);
                                std::wstring formattedEntityName = getFormattedEntityName(entityName);
                                if (formattedEntityName == userInputName) {
                                    BSTR entityHandle;
                                    pAcadEntity->get_Handle(&entityHandle);
                                    std::wstring key = std::wstring((wchar_t*)entityHandle);
                                    workEntitiesMap[key] = pAcadEntity;
                                }
                            }

                            std::vector<std::wstring> outputVec;
                            for (const auto& entity : workEntitiesMap) {
                                std::wstring output;
                                IAcadEntity* pAcadEntity = entity.second;
                                BSTR entityName;
                                pAcadEntity->get_EntityName(&entityName);
                                std::wstring formattedEntityName = getFormattedEntityName(entityName);
                                output = handleEntity(entity, formattedEntityName);
                                outputVec.push_back(output);
                            }

                            std::wcout << L"Выберите ID примитива, с которым желаете работать: \n";
                            std::wstring userInputID;
                            std::getline(std::wcin, userInputID);

                            // Find the entity in the map based on the user's input
                            auto it = entitiesMap.find(userInputID);
                            if (it != entitiesMap.end()) {
                                IAcadEntity* pAcadEntity = it->second;

                                bool inputAction = false;
                                while (inputAction == false) {
                                    bool isBlock = false;
                                    long entityType;
                                    pAcadEntity->get_EntityType(&entityType);
                                    if (entityType == acBlockReference)
                                        isBlock = true;
                                    if (isBlock == true)
                                        std::wcout << L"Выберите действие: \n\n" << L"\t1. Сохранить в файл" << L"\n\t2. Удалить" << L"\n\t3. Выход" << L"\n\t4. Расшифоврать блок\n";
                                    else
                                        std::wcout << L"Выберите действие: \n\n" << L"\t1. Сохранить в файл" << L"\n\t2. Удалить" << L"\n\t3. Выход\n";
                                    std::wstring userInputAction;
                                    std::getline(std::wcin, userInputAction);
                                    int action = 0;
                                    try {
                                        action = std::stoi(userInputAction);
                                    }
                                    catch (const std::invalid_argument&) {
                                        std::wcout << L"Некорректный ввод. Пожалуйста, введите число.\n";
                                    }
                                    switch (action)
                                    {
                                    case 1:
                                        for (const auto& str : outputVec) {
                                            if (str.find(userInputID) != std::wstring::npos) {
                                                saveStringToFile(str, outputFile);
                                            }
                                            else {
                                            }
                                        }
                                        std::wcout << L"Примитив успешно сохранен в файл\n";
                                        inputAction = true;

                                    case 2:
                                        inputAction = true;
                                        //AcDbObjectId id;
                                        //pAcadEntity->get_ObjectID(id);
                                        //removeEntity(id);

                                    case 3:
                                        inputAction = true;
                                        break;

                                    case 4:
                                        if (isBlock == true) {
                                            //handleBlockReferenceEntity(it);
                                        }
                                        else {
                                            std::wcout << L"Такого действия не существует(Введите цифру 1, 2 или 3)\n";
                                            break;
                                        }
                                    default:
                                        std::wcout << L"Такого действия не существует(Введите цифру 1, 2 или 3)\n";
                                        break;
                                    }
                                }
                                pAcadEntity->Release();
                            }
                            else {
                                std::wcout << "Entity(ID) не найдено" << std::endl;
                            }
                        }

                        // Release the ModelSpace
                        pAcadModelSpace->Release();
                    }


                    // Now it should be safe to close the document
                    VARIANT vSaveChanges;
                    vSaveChanges.vt = VT_BOOL;
                    vSaveChanges.boolVal = VARIANT_TRUE;
                    VARIANT vFilename;
                    vFilename.vt = VT_EMPTY;
                    pAcadDoc->Close(vSaveChanges, vFilename);
                    DebugPrint(_T("DWG file saved and closed"));


                    // Release the document
                    pAcadDoc->Release();
                }

                // Release the Documents collection
                pAcadDocs->Release();
            }

            // Quit AutoCAD
            pAcadApp->Quit();
            DebugPrint(_T("AutoCAD quit"));

            // Release the AutoCAD application
            pAcadApp->Release();
        }

        // Uninitialize COM
        CoUninitialize();
        DebugPrint(_T("COM uninitialized"));
    }

    return 0;
}
