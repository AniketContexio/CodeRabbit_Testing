package com.contexio.platformXPlus.insights.serviceImpl;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.UUID;
import java.util.ArrayList;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Map;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.http.HttpStatus;
//import org.springframework.data.mongodb.core.aggregation.VariableOperators.Map;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.web.multipart.MultipartFile;
//import org.thymeleaf.expression.Arrays;

import com.contexio.platformXPlus.dto.RegistrationDTO;
import com.contexio.platformXPlus.insights.constants.Grading;
import com.contexio.platformXPlus.insights.dto.BestSellerDTO;
import com.contexio.platformXPlus.insights.dto.CriteriaDTO;
import com.contexio.platformXPlus.insights.dto.CustomerResponseDTO;
import com.contexio.platformXPlus.insights.dto.InsightsContentDto;
import com.contexio.platformXPlus.insights.dto.InsightsContentDto.ReportDetail;
import com.contexio.platformXPlus.insights.dto.InsightsProductDataDTO;
import com.contexio.platformXPlus.insights.dto.ParameterScoreDto;
import com.contexio.platformXPlus.insights.dto.ProductDataDTOInsights;
import com.contexio.platformXPlus.insights.dto.RankDTO;
import com.contexio.platformXPlus.insights.dto.RankDetail;
import com.contexio.platformXPlus.insights.repository.EcomXLDashboardRepository;
import com.contexio.platformXPlus.insights.repository.UploadRepository;
import com.contexio.platformXPlus.insights.service.UploadService;
import com.contexio.platformXPlus.productconfig.dto.BatchDetailsDTO;
import com.contexio.platformXPlus.productconfig.dto.ProductDTO;

import jakarta.servlet.http.HttpSession;

@Service
public class UploadServiceImpl implements UploadService {

    @Autowired
    private EcomXLDashboardRepository brandDashboardRepository;

    @Autowired
    private UploadRepository uploadRepository;

    @Autowired
    private MongoTemplate mongoTemplate;

      // Best Seller Data
      @Override
      public ResponseEntity<String> generateBestSellerData(String clientId, String parentCategory,
              String channelId, String brandSubbrandId, String projectId, String crawlDate, String crawlTime, String zipcode, String userName) {
   
          List<String> clientIds = Collections.singletonList(clientId);
          List<String> projectIds = Collections.singletonList(projectId);
          List<String> brandSubbrandIds = Collections.singletonList(brandSubbrandId);
   
          // Step 1: Validate Mappings
          List<String> validClientProjectMappings = brandDashboardRepository.findMappedProjectIdsClientIds(clientIds,
                  projectIds);
            System.out.println("validClientProjectMappings"+validClientProjectMappings);
   
          List<String> validClientSubBrandProjectMappings = brandDashboardRepository.findMappedProjectIdsClientIdsBrandSubbrandId(brandSubbrandIds,
                  validClientProjectMappings);
                  System.out.println("validClientSubBrandProjectMappings"+validClientSubBrandProjectMappings);
          List<String> validClientChannelMappings = brandDashboardRepository.findMappedChannelIdsClientIds(clientIds,
                  channelId);
                  System.out.println("Client and channel Mapping : "+clientIds+channelId);
                  System.out.println("validClientChannelMappings"+validClientChannelMappings);


          if (validClientProjectMappings.isEmpty() || validClientChannelMappings.isEmpty()) {
              return ResponseEntity.badRequest().body("Invalid client-project or client-channel mapping.");
          }
   
          String mappingClientProjectId = validClientProjectMappings.get(0);
          String mappingClientChannelId = validClientChannelMappings.get(0);
          String mappingClientBrandSubbrandId = validClientSubBrandProjectMappings.get(0);
   
          // Fetch Child Categories for the Parent
          List<String> childCategories = brandDashboardRepository.findChildCategoriesByParentCategory(parentCategory,
                  clientId, projectId);
   
          if (childCategories == null || childCategories.isEmpty()) {
              return ResponseEntity.badRequest().body("No child categories found for parent category: " + parentCategory);
          }
   
          String timestamp = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss").format(new Date());
   
          // Step 2: Process each child category and accumulate the total count
          int totalCount = 0;
   
          for (String childCategory : childCategories) {
              int count = uploadRepository.countBestSellerByCategoryAndMapping(
                      childCategory, parentCategory, mappingClientBrandSubbrandId, mappingClientChannelId);
   
              totalCount += count;
   
              BestSellerDTO entity = new BestSellerDTO();
              entity.setParentCategoryId(parentCategory);
              entity.setMapping_client_project(mappingClientBrandSubbrandId);
              entity.setMapping_client_channel(mappingClientChannelId);
              entity.setCrawlDate(crawlDate);
              entity.setCrawlTime(crawlTime);
              entity.setZipcode(zipcode);
              entity.setType("childCategory");
              entity.setType_id(childCategory);
              entity.setCount(count);
              entity.setTimestamp(timestamp);
   
              mongoTemplate.save(entity, "insights_bestSeller");
          }
   
          // Step 3: Save aggregated parent category entry
          BestSellerDTO parentEntity = new BestSellerDTO();
          parentEntity.setMapping_client_project(mappingClientBrandSubbrandId);
          parentEntity.setMapping_client_channel(mappingClientChannelId);
          parentEntity.setCrawlDate(crawlDate);
          parentEntity.setCrawlTime(crawlTime);
          parentEntity.setZipcode(zipcode);
          parentEntity.setType("parentCategory");
          parentEntity.setType_id(parentCategory);
          parentEntity.setCount(totalCount);
          parentEntity.setTimestamp(timestamp);
   
          mongoTemplate.save(parentEntity, "insights_bestSeller");
   
          return ResponseEntity
                  .ok("Best seller scores saved for all child categories and parent category: " + parentCategory);
      }
   // ContentQualityGeneration@Override
   @Override
   public ResponseEntity<String> generateContentQualityData(
           String clientId, String parentCategory,
           String channelId, String brandSubbrandId, String projectId,
           String crawlDate, String crawlTime,
           String zipcode, String userName) {

       List<String> clientIds = Collections.singletonList(clientId);
       List<String> projectIds = Collections.singletonList(projectId);
       List<String> brandSubbrandIds = Collections.singletonList(brandSubbrandId);

       List<String> validClientProjectMappings = brandDashboardRepository.findMappedProjectIdsClientIds(clientIds,
               projectIds);

       List<String> validClientSubBrandProjectMappings = brandDashboardRepository
               .findMappedProjectIdsClientIdsBrandSubbrandId(brandSubbrandIds,
                       validClientProjectMappings);

       List<String> validClientChannelMappings = brandDashboardRepository.findMappedChannelIdsClientIds(clientIds,
               channelId);

       if (validClientProjectMappings.isEmpty() || validClientChannelMappings.isEmpty()) {
           return ResponseEntity.badRequest().body("Invalid client-project or client-channel mapping.");
       }

       String mappingClientProjectId = validClientProjectMappings.get(0);
       String mappingClientChannelId = validClientChannelMappings.get(0);
       String mappingClientBrandSubbrandId = validClientSubBrandProjectMappings.get(0);

       List<String> childCategories = brandDashboardRepository.findChildCategoriesByParentCategory(parentCategory,
               clientId, projectId);
       if (childCategories == null || childCategories.isEmpty()) {
           return ResponseEntity.badRequest().body("No child categories found for parent category: " + parentCategory);
       }

       List<CriteriaDTO> criteriaList = brandDashboardRepository.getActiveCriteria(clientId, projectId);
       if (criteriaList.isEmpty()) {
           return ResponseEntity.badRequest().body("No active criteria found for client/project.");
       }

       Map<String, String> parameterGroupMap = brandDashboardRepository.getParameterGroupMapping(clientId, projectId);

       List<InsightsContentDto> contentList = new ArrayList<>();
       Map<String, List<InsightsContentDto>> childCategoryMap = new HashMap<>();

       for (String childCategory : childCategories) {
           List<ProductDataDTOInsights> productDataList = uploadRepository.findByFilters(
                   mappingClientBrandSubbrandId, mappingClientChannelId, childCategory, parentCategory,
                   crawlDate, crawlTime, zipcode);

           if (productDataList.isEmpty())
               continue;

           List<InsightsContentDto> perProductList = new ArrayList<>();

           for (ProductDataDTOInsights product : productDataList) {
               Map<String, Object> props = product.getProperties();

               InsightsContentDto content = new InsightsContentDto();
               content.setParentCategory(parentCategory);
               content.setMapping_client_project(mappingClientBrandSubbrandId);
               content.setMapping_client_channel(mappingClientChannelId);
               content.setTimestamp(LocalDateTime.now().toString());
               content.setCrawlDate(crawlDate);
               content.setCrawlTime(crawlTime);
               content.setZipcode(zipcode);
               content.setType("product");
               content.setType_id((String) props.get("PRODUCT_ID"));

               List<ReportDetail> details = new ArrayList<>();
               Map<String, List<Integer>> paramScoreMap = new HashMap<>();

               for (CriteriaDTO crit : criteriaList) {
                   String criteriaName = crit.getCriteriaName();
                   String gradingType = crit.getGradingType();
                   Object value = props.get(criteriaName);

                   int score;
                   try {
                       score = switch (gradingType.toLowerCase()) {
                           case "binary" -> Grading.binary(value);
                           case "description" -> Grading.description((String) value);
                           case "features" -> Grading.features(value);
                           case "images" -> Grading.images((String) value);
                           case "specs" -> Grading.specs((String) value);
                           case "videos" -> Grading.videos((String) value);
                           case "reviews" -> Grading.reviews(value);
                           case "rating" -> Grading.rating(value);
                           case "search" -> Grading.searchPosition((Integer) value);
                           default -> 0;
                       };
                   } catch (Exception e) {
                       score = 0;
                   }

                   details.add(new ReportDetail(criteriaName, String.valueOf(score)));

                   String paramGroup = parameterGroupMap.get(criteriaName);
                   if (paramGroup != null) {
                       paramScoreMap.computeIfAbsent(paramGroup, k -> new ArrayList<>()).add(score);
                   }
               }

               List<ParameterScoreDto> parameterScores = paramScoreMap.entrySet().stream()
                       .map(entry -> new ParameterScoreDto(entry.getKey(),
                               (int) Math.round(
                                       entry.getValue().stream().mapToInt(Integer::intValue).average().orElse(0))))
                       .toList();

               content.setDetails(details);
               content.setParameterScores(parameterScores);
               int overallScore = (int) Math.round(parameterScores.stream()
                       .mapToInt(ParameterScoreDto::getAverageScore).average().orElse(0));
               content.setOverallScore(overallScore);

               contentList.add(content);
               perProductList.add(content);
           }

           // Save product list under child category
           childCategoryMap.put(childCategory, perProductList);

           // Compute averages for child category
           Map<String, List<Integer>> criteriaScoreMap = new HashMap<>();
           Map<String, List<Integer>> parameterScoreMap = new HashMap<>();

           for (InsightsContentDto dto : perProductList) {
               for (ReportDetail detail : dto.getDetails()) {
                   criteriaScoreMap.computeIfAbsent(detail.getReport_detail_id(), k -> new ArrayList<>())
                           .add(Integer.parseInt(detail.getReport_detail_score()));
               }
               for (ParameterScoreDto scoreDto : dto.getParameterScores()) {
                   parameterScoreMap.computeIfAbsent(scoreDto.getParameterGroupName(), k -> new ArrayList<>())
                           .add(scoreDto.getAverageScore());
               }
           }

           List<ReportDetail> avgDetails = criteriaScoreMap.entrySet().stream()
                   .map(e -> new ReportDetail(e.getKey(),
                           String.valueOf(
                                   Math.round(e.getValue().stream().mapToInt(Integer::intValue).average().orElse(0)))))
                   .toList();

           List<ParameterScoreDto> avgParamScores = parameterScoreMap.entrySet().stream()
                   .map(e -> new ParameterScoreDto(e.getKey(),
                           (int) Math.round(e.getValue().stream().mapToInt(Integer::intValue).average().orElse(0))))
                   .toList();

           InsightsContentDto childContent = new InsightsContentDto();
           childContent.setParentCategory(parentCategory);
           childContent.setType("childcategory");
           childContent.setType_id(childCategory);
           childContent.setMapping_client_project(mappingClientBrandSubbrandId);
           childContent.setMapping_client_channel(mappingClientChannelId);
           childContent.setTimestamp(LocalDateTime.now().toString());
           childContent.setCrawlDate(crawlDate);
           childContent.setCrawlTime(crawlTime);
           childContent.setZipcode(zipcode);
           childContent.setDetails(avgDetails);
           childContent.setParameterScores(avgParamScores);
           int childOverallScore = (int) Math.round(avgParamScores.stream()
                   .mapToInt(ParameterScoreDto::getAverageScore).average().orElse(0));
           childContent.setOverallScore(childOverallScore);

           contentList.add(childContent);
       }

       // Compute averages for parent category
       Map<String, List<Integer>> parentCriteriaMap = new HashMap<>();
       Map<String, List<Integer>> parentParameterMap = new HashMap<>();

       for (List<InsightsContentDto> childList : childCategoryMap.values()) {
           for (InsightsContentDto dto : childList) {
               for (ReportDetail detail : dto.getDetails()) {
                   parentCriteriaMap.computeIfAbsent(detail.getReport_detail_id(), k -> new ArrayList<>())
                           .add(Integer.parseInt(detail.getReport_detail_score()));
               }
               for (ParameterScoreDto scoreDto : dto.getParameterScores()) {
                   parentParameterMap.computeIfAbsent(scoreDto.getParameterGroupName(), k -> new ArrayList<>())
                           .add(scoreDto.getAverageScore());
               }
           }
       }

       List<ReportDetail> parentAvgDetails = parentCriteriaMap.entrySet().stream()
               .map(e -> new ReportDetail(e.getKey(),
                       String.valueOf(
                               Math.round(e.getValue().stream().mapToInt(Integer::intValue).average().orElse(0)))))
               .toList();

       List<ParameterScoreDto> parentAvgParams = parentParameterMap.entrySet().stream()
               .map(e -> new ParameterScoreDto(e.getKey(),
                       (int) Math.round(e.getValue().stream().mapToInt(Integer::intValue).average().orElse(0))))
               .toList();

       InsightsContentDto parentContent = new InsightsContentDto();
       parentContent.setParentCategory(parentCategory);
       parentContent.setType("parentcategory");
       parentContent.setType_id(parentCategory);
       parentContent.setMapping_client_project(mappingClientBrandSubbrandId);
       parentContent.setMapping_client_channel(mappingClientChannelId);
       parentContent.setTimestamp(LocalDateTime.now().toString());
       parentContent.setCrawlDate(crawlDate);
       parentContent.setCrawlTime(crawlTime);
       parentContent.setZipcode(zipcode);
       parentContent.setDetails(parentAvgDetails);
       parentContent.setParameterScores(parentAvgParams);
       int parentOverallScore = (int) Math.round(parentAvgParams.stream()
               .mapToInt(ParameterScoreDto::getAverageScore).average().orElse(0));
       parentContent.setOverallScore(parentOverallScore);

       contentList.add(parentContent);

       // Save all to Mongo
       for (InsightsContentDto content : contentList) {
           mongoTemplate.save(content);
       }

       return ResponseEntity.ok("Successfully uploaded " + contentList.size() + " content quality entries.");
   }

   // Rank Upload Generation
   @Override
   public ResponseEntity<String> uploadRankData(MultipartFile uploadFile, String clientId, String childCategory,
           String parentCategory, String channelId, String brandSubbrandId, String projectId, String crawlDate,
           String crawlTime,
           String zipcode, String searchKey, String userName) {

       List<String> clientIds = Collections.singletonList(clientId);
       List<String> projectIds = Collections.singletonList(projectId);
       List<String> brandSubbrandIds = Collections.singletonList(brandSubbrandId);

       List<String> validClientProjectMappings = brandDashboardRepository.findMappedProjectIdsClientIds(clientIds,
               projectIds);
       List<String> validClientSubBrandProjectMappings = brandDashboardRepository
               .findMappedProjectIdsClientIdsBrandSubbrandId(brandSubbrandIds,
                       validClientProjectMappings);
       List<String> validClientChannelMappings = brandDashboardRepository.findMappedChannelIdsClientIds(clientIds,
               channelId);

       try (InputStream is = uploadFile.getInputStream(); Workbook workbook = new XSSFWorkbook(is)) {
           Sheet sheet = workbook.getSheetAt(0);
           Map<String, RankDTO> rankMap = new HashMap<>();

           for (int i = 1; i <= sheet.getLastRowNum(); i++) {
               Row row = sheet.getRow(i);
               if (row == null)
                   continue;

               String keyword = row.getCell(1).getStringCellValue().trim();
               String keywordId = row.getCell(0).getStringCellValue().trim();
               int rank = (int) row.getCell(4).getNumericCellValue();
               String brand = row.getCell(5).getStringCellValue().trim();
               String channel = row.getCell(6).getStringCellValue().trim();
               String title = row.getCell(7).getStringCellValue().trim();
               String URL = row.getCell(8).getStringCellValue().trim();
               int price = (int) row.getCell(9).getNumericCellValue();
               String imageURL = row.getCell(10).getStringCellValue().trim();
               String skuID = row.getCell(11).getStringCellValue().trim();

               RankDTO doc = rankMap.computeIfAbsent(keyword, k -> {
                   RankDTO r = new RankDTO();
                   r.setKeyword(keyword);
                   r.setKeyword_id(keywordId);
                   r.setChildCategory(childCategory);
                   r.setParentCategory(parentCategory);
                   r.setMapping_client_project(validClientSubBrandProjectMappings.get(0));
                   r.setMapping_client_channel(validClientChannelMappings.get(0));
                   r.setCrawlDate(crawlDate);
                   r.setCrawlTime(crawlTime);
                   r.setZipcode(zipcode);
                   r.setTimestamp(LocalDateTime.now().toString());
                   r.setSearchKey(searchKey);
                   r.setRankDetails(new ArrayList<>());
                   return r;
               });

               RankDetail detail = new RankDetail();
               detail.setRank(rank);
               detail.setBrand(brand);
               detail.setChannel(channel);
               detail.setTitle(title);
               detail.setUrl(URL);
               detail.setPrice(price);
               detail.setImage_url(imageURL);
               detail.setSku_id(skuID);

               doc.getRankDetails().add(detail);
           }

           for (RankDTO doc : rankMap.values()) {
               mongoTemplate.save(doc, "insights_rank");
           }

           return ResponseEntity.ok("Rank data uploaded successfully.");

       } catch (Exception e) {
           e.printStackTrace();
           return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                   .body("Failed to upload rank data: " + e.getMessage());
       }
   }

   @Override
   public ResponseEntity<String> generateCustomerResponseData(String clientId, String parentCategory, String channelId,
           String brandSubbrandId,
           String projectId, String crawlDate, String crawlTime, String zipcode, String userName) {

       List<String> clientIds = Collections.singletonList(clientId);
       List<String> projectIds = Collections.singletonList(projectId);
       List<String> brandSubbrandIds = Collections.singletonList(brandSubbrandId);

       String keywordData = "";

       List<String> validClientProjectMappings = brandDashboardRepository.findMappedProjectIdsClientIds(clientIds,
               projectIds);
       List<String> validClientSubBrandProjectMappings = brandDashboardRepository
               .findMappedProjectIdsClientIdsBrandSubbrandId(brandSubbrandIds,
                       validClientProjectMappings);
       List<String> validClientChannelMappings = brandDashboardRepository.findMappedChannelIdsClientIds(clientIds,
               channelId);

       if (validClientProjectMappings.isEmpty() || validClientChannelMappings.isEmpty()) {
           return ResponseEntity.badRequest().body("Invalid client-project or client-channel mapping.");
       }

       String mappingClientProjectId = validClientProjectMappings.get(0);
       String mappingClientChannelId = validClientChannelMappings.get(0);
       String mappingClientBrandSubbrandId = validClientSubBrandProjectMappings.get(0);

       String timestamp = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss").format(new Date());

       List<String> childCategories = brandDashboardRepository.findChildCategoriesByParentCategory(
               parentCategory, clientId, projectId);

       if (childCategories == null || childCategories.isEmpty()) {
           return ResponseEntity.badRequest().body("No child categories found for parent category: " + parentCategory);
       }

       List<Double> parentPositiveList = new ArrayList<>();
       List<Double> parentNegativeList = new ArrayList<>();
       List<Double> parentNeutralList = new ArrayList<>();
       List<Double> childRatingList = new ArrayList<>();
       List<Double> parentRatingList = new ArrayList<>();

       for (String childCategory : childCategories) {

           List<ProductDataDTOInsights> productDataList = uploadRepository.findByFilters(
                   mappingClientBrandSubbrandId, mappingClientChannelId, childCategory, parentCategory,
                   crawlDate, crawlTime, zipcode);
           System.out.println("ProductDataList: " + productDataList);

           List<Double> childPositiveList = new ArrayList<>();
           List<Double> childNegativeList = new ArrayList<>();
           List<Double> childNeutralList = new ArrayList<>();

           for (ProductDataDTOInsights product : productDataList) {
               String reviews = product.getReview();
               System.out.println("Reviews: " + reviews);
               if (reviews == null || reviews.isEmpty())
                   continue;
               try {
                   keywordData = reviews;
                   double[] scores = getSentimentForReview(reviews);
                   System.out.println("Scores: " + Arrays.toString(scores));
                   Double rating = (Double) product.getProperties().get("AVG. STAR RATINGS");
                   if (rating != null) {
                       childRatingList.add(rating);
                   }

                   // Save review as product-level entry
                   CustomerResponseDTO productReviewEntry = new CustomerResponseDTO();
                   productReviewEntry.setType("product");
                   productReviewEntry.setType_id(product.getProductId());
                   productReviewEntry.setParentCategory(parentCategory);
                   productReviewEntry.setChildCategory(childCategory);
                   productReviewEntry.setMapping_client_project(mappingClientBrandSubbrandId);
                   productReviewEntry.setMapping_client_channel(mappingClientChannelId);
                   productReviewEntry.setTimestamp(timestamp);
                   productReviewEntry.setCrawlDate(crawlDate);
                   productReviewEntry.setCrawlTime(crawlTime);
                   productReviewEntry.setZipcode(zipcode);
                   productReviewEntry.setKeywords(reviews);
                   productReviewEntry.setPositive_score((int) Math.round(scores[0]));
                   productReviewEntry.setNegative_score((int) Math.round(scores[1]));
                   productReviewEntry.setNeutral_score((int) Math.round(scores[2]));
                   productReviewEntry.setRating_score((Double) product.getProperties().get("AVG. STAR RATINGS"));

                   mongoTemplate.save(productReviewEntry, "insights_customer_response");

                   // Collect scores for child category aggregation
                   childPositiveList.add(scores[0]);
                   childNegativeList.add(scores[1]);
                   childNeutralList.add(scores[2]);

               } catch (IOException e) {
                   System.err.println("Sentiment analysis failed: " + e.getMessage());
               }

           }

           // Save one entry per childCategory
           if (!childPositiveList.isEmpty()) {
               double avgChildPositive = childPositiveList.stream().mapToDouble(Double::doubleValue).average()
                       .orElse(0);
               double avgChildNegative = childNegativeList.stream().mapToDouble(Double::doubleValue).average()
                       .orElse(0);
               double avgChildNeutral = childNeutralList.stream().mapToDouble(Double::doubleValue).average().orElse(0);
               double avgChildRating = childRatingList.stream().mapToDouble(Double::doubleValue).average().orElse(0);

               CustomerResponseDTO childEntry = new CustomerResponseDTO();
               childEntry.setType("childCategory");
               childEntry.setType_id(childCategory);
               childEntry.setParentCategory(parentCategory);
               childEntry.setChildCategory(childCategory);
               childEntry.setMapping_client_project(mappingClientBrandSubbrandId);
               childEntry.setMapping_client_channel(mappingClientChannelId);
               childEntry.setTimestamp(timestamp);
               childEntry.setCrawlDate(crawlDate);
               childEntry.setCrawlTime(crawlTime);
               childEntry.setZipcode(zipcode);
               childEntry.setKeywords(keywordData);
               childEntry.setPositive_score((int) Math.round(avgChildPositive));
               childEntry.setNegative_score((int) Math.round(avgChildNegative));
               childEntry.setNeutral_score((int) Math.round(avgChildNeutral));
               childEntry.setRating_score(avgChildRating);

               mongoTemplate.save(childEntry, "insights_customer_response");

               // Add to parent-level aggregation
               parentPositiveList.add(avgChildPositive);
               parentNegativeList.add(avgChildNegative);
               parentNeutralList.add(avgChildNeutral);
               parentRatingList.add(avgChildRating);

           }
       }

       // Save one entry for parentCategory
       if (!parentPositiveList.isEmpty()) {
           double avgParentPositive = parentPositiveList.stream().mapToDouble(Double::doubleValue).average().orElse(0);
           double avgParentNegative = parentNegativeList.stream().mapToDouble(Double::doubleValue).average().orElse(0);
           double avgParentNeutral = parentNeutralList.stream().mapToDouble(Double::doubleValue).average().orElse(0);
           double avgParentRating = parentRatingList.stream().mapToDouble(Double::doubleValue).average().orElse(0);

           CustomerResponseDTO parentEntry = new CustomerResponseDTO();
           parentEntry.setType("parentCategory");
           parentEntry.setType_id(parentCategory);
           parentEntry.setParentCategory(parentCategory);
           parentEntry.setChildCategory("");
           parentEntry.setMapping_client_project(mappingClientBrandSubbrandId);
           parentEntry.setMapping_client_channel(mappingClientChannelId);
           parentEntry.setTimestamp(timestamp);
           parentEntry.setCrawlDate(crawlDate);
           parentEntry.setCrawlTime(crawlTime);
           parentEntry.setZipcode(zipcode);
           parentEntry.setKeywords(keywordData);
           parentEntry.setPositive_score((int) Math.round(avgParentPositive));
           parentEntry.setNegative_score((int) Math.round(avgParentNegative));
           parentEntry.setNeutral_score((int) Math.round(avgParentNeutral));
           parentEntry.setRating_score(avgParentRating);

           mongoTemplate.save(parentEntry, "insights_customer_response");
       }

       return ResponseEntity.ok("Product, childCategory, and parentCategory sentiment entries stored successfully.");
   }

   private double[] getSentimentForReview(String review) throws IOException {
       System.out.println("Reached Here....");
       System.out.println("Review: " + review);

       // Escape quotes
       review = review.replace("\"", "\\\"");
       System.out.println("Escaped review: " + review);

       String[] cmd = new String[] { "cmd.exe", "/c", "py D:\\sentiment_analysis_percentage.py \"" + review + "\"" };
       ProcessBuilder pb = new ProcessBuilder(cmd);
       pb.redirectErrorStream(true); // Combine stdout and stderr

       Process process = pb.start();

       BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
       String line;
       String sentimentLine = null;
       while ((line = reader.readLine()) != null) {
           System.out.println("PYTHON OUT: " + line); // Show all lines for debugging
           if (line.matches("\\d+(\\.\\d+)?\\s*,\\s*\\d+(\\.\\d+)?\\s*,\\s*\\d+(\\.\\d+)?")) {
               sentimentLine = line; // Extract the actual sentiment values
           }
       }

       if (sentimentLine != null) {
           String[] parts = sentimentLine.split(",");
           double[] scores = new double[3];
           for (int i = 0; i < 3; i++) {
               scores[i] = Double.parseDouble(parts[i].trim());
           }
           return scores;
       }

       // fallback if parsing fails
       return new double[] { 0.0, 0.0, 0.0 };
   }


    // Added by Aniket Vishwakarma for Fetch Channel on 24/07/2025
    @Override
    public List<Map<String, String>> fetchChannelData(String mappingId) {
        return uploadRepository.fetchChannelData(mappingId);
    }

    // Added by Aniket Vishwakarma for Fetch Product Data on 24/07/2025
    @Override
    public List<Map<String, String>> fetchProductDataInsights(String clientName, String subBrand, String projectName,
            String parentCategory, String channelName, String filterDate) {
        List<ProductDataDTOInsights> productDataIds = uploadRepository.fetchProductDataInsights(clientName, subBrand, projectName,
                parentCategory, channelName, filterDate);

            List<Map<String, String>> enrichedList = new ArrayList<>();

            for (ProductDataDTOInsights dto : productDataIds) {
                Map<String, String> map = new HashMap<>();

            // map.put("crawlDate", dto.getCrawlDate());

                // Map<String, String> subBrandMap = uploadRepository.getSubBrandDetails(subBrand);
                // Map<String, String> projectMap = uploadRepository.getProjectDetails(dto.getMapping_client_project());
                Map<String, String> categoryMap = uploadRepository.getCategoryDetails(dto.getParentCategory());
                Map<String, String> channelMap = uploadRepository.getChannelDetails(dto.getMapping_client_channel());

                // if (subBrandMap != null) map.putAll(subBrandMap);
                // if (projectMap != null) map.putAll(projectMap);
                if (categoryMap != null) map.putAll(categoryMap);
                if (channelMap != null) map.putAll(channelMap);

                // Also include mapping IDs from Mongo directly
                map.put("date", dto.getPreProcessDate());
                map.put("crawlTime", (String)dto.getProperties().get("crawlTime"));
                map.put("crawlDate", (String)dto.getProperties().get("crawlDate"));
                map.put("mappingSubBrandId", subBrand);
                map.put("mappingBrandSubbrandProjectId", dto.getMapping_client_project());
                map.put("mappingClientChannelId", dto.getMapping_client_channel());

                enrichedList.add(map);
            }

            return enrichedList;
    }

   @Override
public boolean productUploadInsights(MultipartFile file, String clientId, String subbrandId,
    String projectId, String channelId, String categoryFrom, String columnNumber,
    String attributeRow, String dataRow, String batchName, String parentCategoryId) throws Exception {

    HttpSession session = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes())
            .getRequest().getSession();
    List<RegistrationDTO> sessionData = (List<RegistrationDTO>) session.getAttribute("USER_SESSION_DATA");
    String custId = sessionData.get(0).getCustomerId();

    String firstName = sessionData.get(0).getFirstName().trim().toLowerCase();
    String lastName = sessionData.get(0).getLastName();

    String UserName = (lastName != null && !lastName.trim().isEmpty())
            ? firstName + "." + lastName.trim().toLowerCase()
            : firstName;

    XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
    XSSFSheet sheet = workbook.getSheetAt(0);

    int attrRowNum = Integer.parseInt(attributeRow);
    int dataRowNum = Integer.parseInt(dataRow);
    int colNum = Integer.parseInt(columnNumber);

    XSSFRow headerRow = sheet.getRow(attrRowNum);
    if (headerRow == null) {
        throw new Exception("Attribute row not found in the file.");
    }

    Query batchQuery = new Query();
    batchQuery.addCriteria(Criteria.where("batchName").is(batchName));
    BatchDetailsDTO existingBatch = mongoTemplate.findOne(batchQuery, BatchDetailsDTO.class, "batch_details");

    if (existingBatch != null) {
        throw new Exception("Batch name '" + batchName + "' already exists. Please choose a different name.");
    }

    String categoryName = null;
    String categoryId = null;

    if (categoryFrom.equalsIgnoreCase("File Name")) {
        String fileName = file.getOriginalFilename();
        fileName = fileName.split("\\.")[0];
        categoryName = fileName;
        categoryId = uploadRepository.findCategoryIdByName(categoryName, clientId, projectId);
        if (categoryId == null) {
            throw new Exception("Category name '" + categoryName + "' from file name not found.");
        }
    } else if (categoryFrom.equalsIgnoreCase("Sheet Name")) {
        categoryName = sheet.getSheetName();
        categoryId = uploadRepository.findCategoryIdByName(categoryName, clientId, projectId);
        if (categoryId == null) {
            throw new Exception("Category name '" + categoryName + "' from sheet name not found.");
        }
    }

    Set<String> categoryNamesSet = new HashSet<>();
    if (categoryFrom.equalsIgnoreCase("Column")) {
        for (int i = dataRowNum; i <= sheet.getLastRowNum(); i++) {
            XSSFRow dataRowObj = sheet.getRow(i);
            if (dataRowObj == null) continue;
            Cell categoryCell = dataRowObj.getCell(colNum);
            if (categoryCell != null && categoryCell.getCellType() != CellType.BLANK) {
                String categoryNameCell = categoryCell.getStringCellValue().trim().replaceAll("[\\u00A0\\r\\n\\t]", "");
                if (!categoryNameCell.isEmpty()) {
                    categoryNamesSet.add(categoryNameCell);
                }
            }
        }
    } else {
        categoryNamesSet.add(categoryName);
    }

    Map<String, String> categoryNameIdMap = new HashMap<>();
    for (String catName : categoryNamesSet) {
        String catId = (categoryFrom.equalsIgnoreCase("Column"))
                ? uploadRepository.findCategoryIdByName(catName, clientId, projectId)
                : categoryId;

        if (catId == null) {
            throw new Exception("Category name '" + catName + "' not found in category master.");
        }

        categoryNameIdMap.put(catName, catId);

        List<String> expectedAttributes = uploadRepository.getAttributesForCategory(catId);
        List<String> excelHeaders = new ArrayList<>();

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            if (categoryFrom.equalsIgnoreCase("Column") && i == colNum) continue;
            Cell cell = headerRow.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                excelHeaders.add(cell.getStringCellValue().trim());
            }
        }

        List<String> missingMappedAttributes = new ArrayList<>();
        for (String expected : expectedAttributes) {
            String normalizedExpected = expected.trim().replace(" ", "_")
                    .replace("&", "and").replace(".", "@@@")
                    .replaceAll("[\\u00A0\\r\\n\\t]", "").toUpperCase();

            boolean found = excelHeaders.stream().anyMatch(header -> 
                header.trim().replace(" ", "_").replace("&", "and").replace(".", "@@@")
                        .replaceAll("[\\u00A0\\r\\n\\t]", "").toUpperCase().equals(normalizedExpected));

            if (!found) {
                missingMappedAttributes.add(expected);
            }
        }

        if (!missingMappedAttributes.isEmpty()) {
            throw new Exception("Mapped attributes missing in Excel for category '" + catName + "': " +
                    String.join(", ", missingMappedAttributes));
        }
    }

    int status = 1;
    String productType = "internal";
    int brandColumnIndex = -1;

    for (int i = 0; i < headerRow.getLastCellNum(); i++) {
        Cell cell = headerRow.getCell(i);
        if (cell != null && cell.getCellType() != CellType.BLANK) {
            if (cell.getStringCellValue().trim().equalsIgnoreCase("BRAND")) {
                brandColumnIndex = i;
                break;
            }
        }
    }

    if (brandColumnIndex == -1) {
        throw new Exception("Brand column not found in attribute row.");
    }

    BatchDetailsDTO batch = new BatchDetailsDTO();
    batch.setBatchName(batchName);
    batch.setClientId(clientId);
    batch.setProjectId(projectId);
    batch.setStatus("active");
    batch.setInsertDate(Instant.now());
    batch.setUpdateDate(Instant.now());
    batch.setInsertId(UserName);
    batch.setUpdateId(UserName);
    mongoTemplate.save(batch, "batch_details");
    String batchIdCreated = batch.getId();

    String cliProId = uploadRepository.findCliProId(clientId, projectId);
    if (cliProId == null) {
        throw new Exception("cli_pro_id not found for clientId=" + clientId + " and projectId=" + projectId);
    }

    String mappingId = uploadRepository.findMappingIdByBrandSubbrandProject(subbrandId, cliProId);
    if (mappingId == null) {
        throw new Exception("mapping_id not found for subbrandId=" + subbrandId + " and cli_pro_id=" + cliProId);
    }

    String mappingChannelId = uploadRepository.fetchClientChannelMappingId(subbrandId, channelId);
    if (mappingChannelId == null) {
        throw new Exception("mapping_channel_id not found for subbrandId=" + subbrandId + " and channelId=" + channelId);
    }

    for (int i = dataRowNum; i <= sheet.getLastRowNum(); i++) {
        XSSFRow dataRowObj = sheet.getRow(i);
        if (dataRowObj == null) continue;

        String rowCategoryName = (categoryFrom.equalsIgnoreCase("Column"))
                ? Optional.ofNullable(dataRowObj.getCell(colNum)).map(Cell::getStringCellValue).orElse("").trim()
                : categoryName;

        String categoryIdForRow = categoryNameIdMap.get(rowCategoryName);
        if (categoryIdForRow == null) {
            throw new Exception("Unexpected category '" + rowCategoryName + "' in row " + i);
        }

        Cell brandCell = dataRowObj.getCell(brandColumnIndex);
        if (brandCell == null || brandCell.getCellType() == CellType.BLANK) {
            throw new Exception("Brand value is missing in row " + i);
        }

        String productName = brandCell.getStringCellValue().trim().replaceAll("[\\u00A0\\r\\n\\t]", "");
        List<String> breadcrumbPath = uploadRepository.fetchBreadcrumbPathBetween(parentCategoryId, categoryIdForRow);
        Collections.reverse(breadcrumbPath);
        String breadcrumb = String.join(" >> ", breadcrumbPath);

        String productIdForRow = uploadRepository.insertProduct(
            productName, categoryIdForRow, productType, breadcrumb, status, UserName
        );

        Map<String, Object> propertiesMap = new HashMap<>();
        for (int j = 0; j < headerRow.getLastCellNum(); j++) {
            Cell headerCell = headerRow.getCell(j);
            Cell dataCell = dataRowObj.getCell(j);
            if (headerCell == null || headerCell.getCellType() == CellType.BLANK) continue;

            String key = headerCell.getStringCellValue().trim()
                    .replace("\u00A0", "").replaceAll("[\\r\\n\\t]", "")
                    .replace(".", "@@@").replace("[", "").replace("]", "");
            Object value = null;

            if (dataCell != null) {
                switch (dataCell.getCellType()) {
                    case STRING -> value = dataCell.getStringCellValue().trim().replaceAll("[\\u00A0\\r\\n\\t]", "");
                    case NUMERIC -> {
                        double numericValue = dataCell.getNumericCellValue();
                        value = (numericValue == Math.floor(numericValue))
                                ? (long) numericValue
                                : BigDecimal.valueOf(numericValue);
                    }
                    case BOOLEAN -> value = dataCell.getBooleanCellValue();
                    case FORMULA -> {
                        try {
                            FormulaEvaluator evaluator = dataCell.getSheet().getWorkbook()
                                    .getCreationHelper().createFormulaEvaluator();
                            CellValue cellValue = evaluator.evaluate(dataCell);
                            value = switch (cellValue.getCellType()) {
                                case STRING -> cellValue.getStringValue();
                                case NUMERIC -> cellValue.getNumberValue();
                                case BOOLEAN -> cellValue.getBooleanValue();
                                default -> null;
                            };
                        } catch (Exception e) {
                            value = null;
                        }
                    }
                }
            }

            if (!key.isEmpty()) {
                propertiesMap.put(key, value);
            }
        }

        ProductDataDTOInsights productDTO = new ProductDataDTOInsights();
        productDTO.setBatchId(batchIdCreated);
        productDTO.setProductId(productIdForRow);
        productDTO.setChildCategoryId(categoryIdForRow);
        productDTO.setParentCategory(parentCategoryId);
        productDTO.setMapping_client_channel(mappingChannelId);
        productDTO.setMapping_client_project(mappingId);
        productDTO.setInsertId(UserName);
        productDTO.setUpdateId(UserName);
        productDTO.setInsertDate(Date.from(LocalDate.now().atStartOfDay(ZoneId.systemDefault()).toInstant()));
        productDTO.setUpdateDate(Date.from(LocalDate.now().atStartOfDay(ZoneId.systemDefault()).toInstant()));
        productDTO.setStage("original_data");
        productDTO.setPreProcessDate(LocalDate.now().toString());
        productDTO.setViewStatus(0);
        productDTO.setCatalogStatus(0);
        productDTO.setComment("");
        productDTO.setProperties(propertiesMap);

        mongoTemplate.save(productDTO, "product_data_insight");
    }

    workbook.close();
    return true;
}


}
