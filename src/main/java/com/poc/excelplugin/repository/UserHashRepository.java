package com.poc.excelplugin.repository;

import com.poc.excelplugin.entity.UserFileHash;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.Optional;

@Repository
public interface UserHashRepository extends JpaRepository<UserFileHash, Long> {

    // Used for Verification: Does this specific hash exist?
    boolean existsByUserIdAndHashKey(String userId, String hashKey);

    // Used for Generation: Check if user already has a file for this entity
    Optional<UserFileHash> findByUserIdAndEntityName(String userId, String entityName);
}