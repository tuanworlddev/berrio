package controllers

import (
	"context"
	"crypto/rand"
	"encoding/hex"
	"errors"
	"fmt"
	"net/http"
	"sync"
	"time"

	"github.com/gin-gonic/gin"
	"omnituan.online/models"
	"omnituan.online/services"
)

const (
	reportJobTTL     = 2 * time.Hour
	reportJobTimeout = 2 * time.Hour
)

var reportJobs = &reportJobStore{items: make(map[string]*reportJob)}

type reportJobStore struct {
	mu    sync.RWMutex
	items map[string]*reportJob
}

type reportJob struct {
	ID          string    `json:"id"`
	Status      string    `json:"status"`
	Progress    int       `json:"progress"`
	TotalChunks int       `json:"totalChunks"`
	DoneChunks  int       `json:"doneChunks"`
	CurrentStep string    `json:"currentStep"`
	Error       string    `json:"error,omitempty"`
	DownloadURL string    `json:"downloadUrl,omitempty"`
	CreatedAt   time.Time `json:"createdAt"`
	UpdatedAt   time.Time `json:"updatedAt"`
	ExpiresAt   time.Time `json:"expiresAt"`

	req      ReportRequest
	dateFrom time.Time
	dateTo   time.Time
	cacheKey string
	data     []byte
}

type reportChunk struct {
	from time.Time
	to   time.Time
}

// CreateReportJob starts a background Wildberries report job.
func CreateReportJob(c *gin.Context) {
	req, dateFrom, dateTo, ok := parseReportRequest(c)
	if !ok {
		return
	}

	cacheKey := reportCacheKey(req)
	if cached, ok := reportCache.Load(cacheKey); ok {
		item := cached.(cachedReport)
		if time.Now().Before(item.expiresAt) {
			job := newReportJob(req, dateFrom, dateTo, cacheKey, 1)
			job.Status = "done"
			job.Progress = 100
			job.DoneChunks = 1
			job.CurrentStep = "Báo cáo đã sẵn sàng từ cache"
			job.DownloadURL = fmt.Sprintf("/api/v1/reports/jobs/%s/download", job.ID)
			job.data = append([]byte(nil), item.data...)
			reportJobs.save(job)
			c.JSON(http.StatusAccepted, reportJobResponse(job))
			return
		}
		reportCache.Delete(cacheKey)
	}

	chunks := splitDateRangeByWeek(dateFrom, dateTo)
	job := newReportJob(req, dateFrom, dateTo, cacheKey, len(chunks))
	reportJobs.save(job)

	go runReportJob(job.ID, chunks)

	c.JSON(http.StatusAccepted, reportJobResponse(job))
}

func GetReportJob(c *gin.Context) {
	job, ok := reportJobs.get(c.Param("id"))
	if !ok {
		c.JSON(http.StatusNotFound, gin.H{"error": "Không tìm thấy job báo cáo"})
		return
	}

	c.JSON(http.StatusOK, reportJobResponse(job))
}

func DownloadReportJob(c *gin.Context) {
	job, ok := reportJobs.get(c.Param("id"))
	if !ok {
		c.JSON(http.StatusNotFound, gin.H{"error": "Không tìm thấy job báo cáo"})
		return
	}
	if job.Status != "done" || len(job.data) == 0 {
		c.JSON(http.StatusConflict, gin.H{"error": "Báo cáo chưa sẵn sàng"})
		return
	}

	writeReportExcel(c, job.data, reportFileName(job.req))
}

func runReportJob(jobID string, chunks []reportChunk) {
	job, ok := reportJobs.get(jobID)
	if !ok {
		return
	}

	ctx, cancel := context.WithTimeout(context.Background(), reportJobTimeout)
	defer cancel()

	reportJobs.update(jobID, func(job *reportJob) {
		job.Status = "running"
		job.CurrentStep = "Đang lấy dữ liệu từ Wildberries"
	})

	reports := make([]models.ReportDetails, 0)
	for index, chunk := range chunks {
		step := fmt.Sprintf("Đang lấy tuần %d/%d (%s - %s)", index+1, len(chunks), chunk.from.Format("02/01/2006"), chunk.to.Format("02/01/2006"))
		reportJobs.update(jobID, func(job *reportJob) {
			job.CurrentStep = step
			job.Progress = progressPercent(job.DoneChunks, job.TotalChunks)
		})

		chunkReports, err := services.GetReportDetails(ctx, job.req.APIKey, chunk.from, chunk.to)
		if err != nil {
			reportJobs.fail(jobID, reportJobErrorMessage(err))
			return
		}
		reports = append(reports, chunkReports...)

		reportJobs.update(jobID, func(job *reportJob) {
			job.DoneChunks = index + 1
			job.Progress = progressPercent(job.DoneChunks, job.TotalChunks)
			job.CurrentStep = fmt.Sprintf("Đã lấy %d/%d tuần", job.DoneChunks, job.TotalChunks)
		})
	}

	data, err := buildReportExcel(reports, job.req)
	if err != nil {
		reportJobs.fail(jobID, "Không thể tạo file Excel")
		return
	}

	storeReportCache(job.cacheKey, data)
	reportJobs.update(jobID, func(job *reportJob) {
		job.Status = "done"
		job.Progress = 100
		job.CurrentStep = "Báo cáo đã sẵn sàng"
		job.DownloadURL = fmt.Sprintf("/api/v1/reports/jobs/%s/download", job.ID)
		job.data = append([]byte(nil), data...)
	})
}

func newReportJob(req ReportRequest, dateFrom, dateTo time.Time, cacheKey string, totalChunks int) *reportJob {
	now := time.Now()
	return &reportJob{
		ID:          randomReportJobID(),
		Status:      "queued",
		Progress:    0,
		TotalChunks: totalChunks,
		DoneChunks:  0,
		CurrentStep: "Đang xếp hàng tạo báo cáo",
		CreatedAt:   now,
		UpdatedAt:   now,
		ExpiresAt:   now.Add(reportJobTTL),
		req:         req,
		dateFrom:    dateFrom,
		dateTo:      dateTo,
		cacheKey:    cacheKey,
	}
}

func splitDateRangeByWeek(dateFrom, dateTo time.Time) []reportChunk {
	chunks := make([]reportChunk, 0)
	for cursor := dateFrom; !cursor.After(dateTo); {
		end := cursor.AddDate(0, 0, 6)
		if end.After(dateTo) {
			end = dateTo
		}
		chunks = append(chunks, reportChunk{from: cursor, to: end})
		cursor = end.AddDate(0, 0, 1)
	}
	return chunks
}

func randomReportJobID() string {
	var bytes [12]byte
	if _, err := rand.Read(bytes[:]); err != nil {
		return fmt.Sprintf("%d", time.Now().UnixNano())
	}
	return hex.EncodeToString(bytes[:])
}

func progressPercent(done, total int) int {
	if total <= 0 {
		return 0
	}
	return int(float64(done) / float64(total) * 100)
}

func reportJobErrorMessage(err error) string {
	switch {
	case errors.Is(err, services.ErrReportRateLimited):
		return "Wildberries đang giới hạn tần suất lấy báo cáo. Hệ thống đã chia theo tuần, vui lòng thử lại sau ít phút."
	case errors.Is(err, context.Canceled), errors.Is(err, context.DeadlineExceeded):
		return "Job lấy báo cáo quá lâu hoặc đã bị hủy."
	default:
		return fmt.Sprintf("Không thể lấy báo cáo: %v", err)
	}
}

func reportJobResponse(job *reportJob) gin.H {
	return gin.H{
		"id":          job.ID,
		"status":      job.Status,
		"progress":    job.Progress,
		"totalChunks": job.TotalChunks,
		"doneChunks":  job.DoneChunks,
		"currentStep": job.CurrentStep,
		"error":       job.Error,
		"downloadUrl": job.DownloadURL,
		"createdAt":   job.CreatedAt,
		"updatedAt":   job.UpdatedAt,
		"expiresAt":   job.ExpiresAt,
	}
}

func (store *reportJobStore) save(job *reportJob) {
	store.mu.Lock()
	defer store.mu.Unlock()
	store.cleanupLocked(time.Now())
	store.items[job.ID] = job
}

func (store *reportJobStore) get(id string) (*reportJob, bool) {
	store.mu.RLock()
	defer store.mu.RUnlock()
	job, ok := store.items[id]
	if !ok || time.Now().After(job.ExpiresAt) {
		return nil, false
	}
	return cloneReportJob(job), true
}

func (store *reportJobStore) update(id string, mutate func(*reportJob)) {
	store.mu.Lock()
	defer store.mu.Unlock()
	job, ok := store.items[id]
	if !ok {
		return
	}
	mutate(job)
	job.UpdatedAt = time.Now()
}

func (store *reportJobStore) fail(id string, message string) {
	store.update(id, func(job *reportJob) {
		job.Status = "failed"
		job.Error = message
		job.CurrentStep = message
	})
}

func (store *reportJobStore) cleanupLocked(now time.Time) {
	for id, job := range store.items {
		if now.After(job.ExpiresAt) {
			delete(store.items, id)
		}
	}
}

func cloneReportJob(job *reportJob) *reportJob {
	clone := *job
	if len(job.data) > 0 {
		clone.data = append([]byte(nil), job.data...)
	}
	return &clone
}
