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

    // This stores the 24-char signature (e.g., af/A/Sg5xTRInoLoDfkdjwp6)
    @Column(name = "hash_key", nullable = false, length = 50)
    private String hashKey;

    @Column(name = "created_at")
    private LocalDateTime createdAt = LocalDateTime.now();
}