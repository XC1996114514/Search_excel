package com.example.excelfileviewer.controller;

import com.example.excelfileviewer.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import jakarta.annotation.PostConstruct;

import java.io.IOException;
import java.util.List;
import java.util.Map;

@Controller
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @RequestMapping(value = "/search", method = {RequestMethod.POST, RequestMethod.GET})
    public String search(@RequestParam String keyword,
                         @RequestParam(defaultValue = "1") int page,
                         @RequestParam(defaultValue = "10") int size,
                         Model model) {
        long startTime = System.currentTimeMillis();

        List<Map<String, Object>> results = excelService.search(keyword);
        int totalResults = results.size();
        int totalPages = (int) Math.ceil((double) totalResults / size);

        int fromIndex = (page - 1) * size;
        int toIndex = Math.min(fromIndex + size, totalResults);
        List<Map<String, Object>> paginatedResults = results.subList(fromIndex, toIndex);

        long endTime = System.currentTimeMillis();
        long searchTime = endTime - startTime;

        model.addAttribute("results", paginatedResults);
        model.addAttribute("keyword", keyword);
        model.addAttribute("currentPage", page);
        model.addAttribute("totalPages", totalPages);
        model.addAttribute("totalResults", totalResults);
        model.addAttribute("pageSize", size);
        model.addAttribute("searchTime", searchTime);

        return "results";
    }

    @PostConstruct
    public void init() throws IOException {
        String folderPath = "excel"; // your excel path
        excelService.loadExcelFiles(folderPath);
    }
}




