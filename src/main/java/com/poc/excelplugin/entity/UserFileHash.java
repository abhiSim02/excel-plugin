package com.poc.excelplugin.entity;

import jakarta.persistence.*;
import lombok.Data;
import lombok.NoArgsConstructor;
import java.time.LocalDateTime;

@Entity
@Table(name = "user_file_hashes")
@Data
@NoArgsConstructor
public class UserFileHash {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "user_id", nullable = false)
    private String userId;

    // New Column: To track which entity this hash belongs to
    @Column(name = "entity_name", nullable = false)
    private String entityName;

    @Column(name = "hash_key", nullable = false, length = 50)
    private String hashKey;

    @Column(name = "created_at", nullable = false, updatable = false)
    private LocalDateTime createdAt;

    // New Column: To track when the hash was last updated
    @Column(name = "updated_at")
    private LocalDateTime updatedAt;

    // Automatically set dates on insert
    @PrePersist
    protected void onCreate() {
        createdAt = LocalDateTime.now();
        updatedAt = LocalDateTime.now();
    }

    // Automatically update timestamp on update
    @PreUpdate
    protected void onUpdate() {
        updatedAt = LocalDateTime.now();
    }
}