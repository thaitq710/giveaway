package vn.goldenboy.giveawayserver.controller;

import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import vn.goldenboy.giveawayserver.model.ApiResponse;
import vn.goldenboy.giveawayserver.service.ExcelService;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/excel")
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService excelService;

    @PostMapping("/process")
    public ResponseEntity<ApiResponse<List<Map<String, Object>>>> processExcel(
            @RequestParam("file") MultipartFile file
    ) {
        List<Map<String, Object>> result = excelService.processExcelFile(file);
        return ResponseEntity.ok(ApiResponse.success(result));
    }
}
